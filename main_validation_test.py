"""
Bulk validation runner for Financial Model Builder.

Usage:
    python main_validation_test.py
"""

import time
import json
import os
import sys
import logging
from pathlib import Path

from sec_fetcher import (
    search_company,
    get_company_facts,
    get_fiscal_years,
    extract_financial_data,
    compute_ltm,
    get_fiscal_year_end_dates,
    get_historical_closing_prices,
    get_current_price,
    get_treasury_yield,
    get_kroll_erp,
    get_industry_peers,
)
from excel import create_excel, _check_status


CONFIG_PATH = Path(__file__).parent / 'config' / 'main_config.json'
TICKERS_CONFIG_PATH = Path(__file__).parent / 'config' / 'main_bulk_test_tickers.json'


def _load_config(path: Path) -> dict:
    with open(path, 'r', encoding='utf-8') as f:
        return json.load(f)


CONFIG = _load_config(CONFIG_PATH)
TICKERS_CONFIG = _load_config(TICKERS_CONFIG_PATH)
BULK_TEST_TICKERS = TICKERS_CONFIG['bulk_test_tickers']
BULK_TEST_SLEEP_SECONDS = CONFIG['defaults']['bulk_test_sleep_seconds']
RISK_FREE_RATE_DEFAULT = CONFIG['defaults']['risk_free_rate']
EQUITY_RISK_PREMIUM_DEFAULT = CONFIG['defaults']['equity_risk_premium']
IMPLIED_COD_DEFAULT = CONFIG['defaults']['implied_cost_of_debt']
LOGGER = logging.getLogger(__name__)


def _safe_name(name: str) -> str:
    return (name
            .replace(' ', '_')
            .replace('/', '-')
            .replace('\\', '-')
            .replace('.', '')
            .replace(',', ''))


def build_model(company_name: str, skip_prices: bool = False, price_date: str = None):
    """Run the full pipeline for one company. Returns (company_info, validation_results) or None on failure."""

    try:
        company_info = search_company(company_name)
    except Exception as e:
        LOGGER.warning("[SKIP] %s", e)
        return None

    print(f"  [OK] {company_info['name']} ({company_info['ticker']}), CIK {company_info['cik']}")

    facts = get_company_facts(company_info['cik'])
    if not facts:
        print(f"  [SKIP] Could not fetch XBRL data for {company_info['ticker']}.")
        return None

    years = get_fiscal_years(facts, n_years=4)
    if not years:
        print(f"  [SKIP] No annual data found for {company_info['ticker']}.")
        return None

    years_display = ', '.join(f'FY{y}' for y in reversed(years))
    print(f"  [OK] Found data for: {years_display}")

    financial_data = extract_financial_data(facts, years, ticker=company_info['ticker'])

    ltm_year = compute_ltm(facts, financial_data, ticker=company_info['ticker'])
    if ltm_year:
        years = financial_data['years']
        ltm_info = financial_data['ltm_info']
        print(f"  [OK] LTM data added: {ltm_info['ann_label']}")

    if not skip_prices:
        fy_end_dates = get_fiscal_year_end_dates(facts, years)
        closing_prices = get_historical_closing_prices(company_info['ticker'], fy_end_dates)
    else:
        closing_prices = {yr: None for yr in years}

    market_cap_data = {}
    shares = financial_data['income_statement']['shares_diluted']
    for yr in years:
        price = closing_prices.get(yr)
        shr = shares.get(yr)
        if price is not None and shr is not None:
            market_cap_data[yr] = price * shr
        else:
            market_cap_data[yr] = None

    financial_data['market_cap'] = market_cap_data
    financial_data['stock_prices'] = closing_prices

    if not skip_prices:
        treasury_yield = get_treasury_yield()
        kroll_erp = get_kroll_erp()
    else:
        treasury_yield = None
        kroll_erp = None

    rf = treasury_yield or RISK_FREE_RATE_DEFAULT
    erp = kroll_erp or EQUITY_RISK_PREMIUM_DEFAULT

    comp_data = []
    if not skip_prices and price_date is None:
        print()
        price_date = input("  Share price date for WACC (YYYY-MM-DD, or Enter for latest): ").strip() or None
    if not skip_prices:
        comp_data, industry_name = get_industry_peers(company_info['ticker'])
        if comp_data:
            print(f"  Comparable companies ({industry_name}):")
            for cd in comp_data:
                print(f"    {cd['name']} ({cd['ticker']}): beta={cd['beta']:.2f}, ${cd['price']:.2f}")

    if not skip_prices:
        current_price_data = get_current_price(company_info['ticker'], price_date)
    else:
        current_price_data = {'price': None, 'date': None}

    latest = years[0]
    inc = financial_data['income_statement']
    bs = financial_data['balance_sheet']
    basic_sh = (inc['shares_basic'].get(latest) or 0) / 1e6
    diluted_sh = (inc['shares_diluted'].get(latest) or 0) / 1e6
    int_exp = abs(inc['interest_expense'].get(latest) or 0)
    total_debt_raw = (bs['st_debt'].get(latest) or 0) + (bs['lt_debt'].get(latest) or 0)
    implied_cod = round(int_exp / total_debt_raw, 4) if total_debt_raw > 0 else IMPLIED_COD_DEFAULT

    financial_data['wacc_inputs'] = {
        'current_price': current_price_data,
        'treasury_yield': rf,
        'kroll_erp': erp,
        'comparables': comp_data,
        'shares_breakdown': {
            'basic': round(basic_sh, 2),
            'rsus': round(max(diluted_sh - basic_sh, 0), 2),
            'options': 0,
            'conv_debt': 0,
            'conv_pref': 0,
        },
        'implied_cod': implied_cod,
    }

    safe = _safe_name(company_info['name'])
    output_dir = os.path.join(os.path.dirname(os.path.abspath(__file__)), 'models')
    os.makedirs(output_dir, exist_ok=True)
    output_file = os.path.join(output_dir, f"{safe}_Financial_Model.xlsx")

    def _stringify_keys(obj):
        if isinstance(obj, dict):
            return {str(k): _stringify_keys(v) for k, v in obj.items()}
        return obj

    tenk_dir = os.path.join(output_dir, '10k')
    os.makedirs(tenk_dir, exist_ok=True)
    tenk_file = os.path.join(tenk_dir, f"{safe}_10k.json")
    with open(tenk_file, 'w', encoding='utf-8') as f:
        json.dump({
            'company': company_info,
            'fiscal_years': years,
            'financial_data': _stringify_keys(financial_data),
        }, f, indent=2)

    validation_results = create_excel(company_info, financial_data, output_file)
    return company_info, validation_results

def run_bulk_validation_test():
    """Run the model through configured companies and print a validation report."""
    print()
    print("=" * 70)
    print("  BULK VALIDATION TEST  --20 Largest Public Companies")
    print("=" * 70)
    print()

    results = []
    for i, ticker in enumerate(BULK_TEST_TICKERS, 1):
        print(f"\n{'-' * 70}")
        print(f"  [{i}/{len(BULK_TEST_TICKERS)}]  Processing: {ticker}")
        print(f"{'-' * 70}")

        try:
            result = build_model(ticker, skip_prices=True)
            if result is None:
                results.append({'ticker': ticker, 'status': 'SKIP', 'reason': 'not found / no data'})
            else:
                company_info, val = result
                results.append({
                    'ticker': ticker,
                    'name': company_info['name'],
                    'status': 'OK',
                    'total': val['total'],
                    'passed': val['passed'],
                    'failed': val['failed'],
                    'checks': val['checks'],
                })
        except Exception as e:
            results.append({'ticker': ticker, 'status': 'ERROR', 'reason': str(e)})
            LOGGER.error("[ERROR] %s: %s", ticker, e)

        # Respect SEC rate limit (10 req/sec, but be conservative)
        time.sleep(BULK_TEST_SLEEP_SECONDS)

    print()
    print()
    print("=" * 70)
    print("  BULK VALIDATION REPORT")
    print("=" * 70)
    print()

    grand_total = 0
    grand_passed = 0
    grand_failed = 0

    for r in results:
        if r['status'] == 'OK':
            grand_total += r['total']
            grand_passed += r['passed']
            grand_failed += r['failed']

            if r['failed'] == 0:
                print(f"  {r['ticker']:>6s}:  {r['passed']}/{r['total']} checks passed  - {r['name']}")
            else:
                fail_details = []
                years = r['checks'][0]['expected'].keys()
                for chk in r['checks']:
                    for yr in years:
                        exp = chk['expected'].get(yr)
                        der = chk['derived'].get(yr)
                        tol_ref_d = chk.get('tolerance_ref')
                        tol_r = tol_ref_d.get(yr, exp) if tol_ref_d else exp
                        passed, diff = _check_status(
                            exp,
                            der,
                            tol_r,
                            tolerance_pct=chk.get('tolerance_pct'),
                        )
                        if not passed:
                            fail_details.append(f"{chk['name']} (FY{yr}, {diff:+.1f})")
                print(f"  {r['ticker']:>6s}:  {r['passed']}/{r['total']} checks passed  - {r['name']}")
                for fd in fail_details:
                    print(f"           FAIL: {fd}")
        elif r['status'] == 'SKIP':
            print(f"  {r['ticker']:>6s}:  SKIPPED - {r['reason']}")
        else:
            print(f"  {r['ticker']:>6s}:  ERROR   - {r['reason']}")

    print()
    print(f"  {'-' * 50}")
    if grand_total > 0:
        pct = grand_passed / grand_total * 100
        print(f"  TOTAL:  {grand_passed}/{grand_total} checks passed  ({pct:.1f}%)")
        print(f"  FAILED: {grand_failed}")
    else:
        print("  No checks were run.")
    print()


if __name__ == '__main__':
    logging.basicConfig(level=logging.INFO, format='%(levelname)s: %(message)s')
    try:
        run_bulk_validation_test()
    except Exception:
        LOGGER.exception("Unhandled error in validation runner")
        sys.exit(1)
