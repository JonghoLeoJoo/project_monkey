"""
Financial Model Builder
=======================
Fetches the latest 4 years of 10-K data from SEC EDGAR and generates
an Excel workbook with:
  - Sheet 1: Historical 3-Statement Model (Income Statement, Balance Sheet, Cash Flow)
  - Sheet 2: DCF Valuation (5-year projection + terminal value)
  - Sheet 3: Data Validation (12 cross-checks with PASS/FAIL status)

Usage:
    python main.py
    python main.py "Apple"
    python main.py AAPL
    python main.py --bulk-test          # Run validation on 20 largest companies
"""

import sys
import os
import json
import time

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
from excel_builder import create_excel, _check_status


# 20 largest US-listed companies by market cap (diverse sectors)
BULK_TEST_TICKERS = [
    'AAPL', 'MSFT', 'NVDA', 'GOOG', 'AMZN',
    'META', 'BRK-B', 'LLY', 'AVGO', 'JPM',
    'WMT', 'V', 'UNH', 'XOM', 'MA',
    'COST', 'PG', 'JNJ', 'HD', 'TSM',
]


def _safe_name(name: str) -> str:
    return (name
            .replace(' ', '_')
            .replace('/', '-')
            .replace('\\', '-')
            .replace('.', '')
            .replace(',', ''))


def build_model(company_name: str, skip_prices: bool = False, auto_select: bool = False, price_date: str = None):
    """Run the full pipeline for one company.  Returns (company_info, validation_results) or None on failure."""

    # Step 1: Find company CIK
    company_info = search_company(company_name, auto_select=auto_select)
    if not company_info:
        print(f"  [SKIP] Could not find '{company_name}' on SEC EDGAR.")
        return None

    print(f"  [OK] {company_info['name']} ({company_info['ticker']}), CIK {company_info['cik']}")

    # Step 2: Download XBRL facts
    facts = get_company_facts(company_info['cik'])
    if not facts:
        print(f"  [SKIP] Could not fetch XBRL data for {company_info['ticker']}.")
        return None

    # Determine the 4 most recent fiscal years
    years = get_fiscal_years(facts, n_years=4)
    if not years:
        print(f"  [SKIP] No annual data found for {company_info['ticker']}.")
        return None

    years_display = ', '.join(f'FY{y}' for y in reversed(years))
    print(f"  [OK] Found data for: {years_display}")

    # Step 3: Extract financial data
    financial_data = extract_financial_data(facts, years, ticker=company_info['ticker'])

    # Step 3b: Try to add quarterly / LTM data
    ltm_year = compute_ltm(facts, financial_data, ticker=company_info['ticker'])
    if ltm_year:
        years = financial_data['years']
        ltm_info = financial_data['ltm_info']
        print(f"  [OK] LTM data added: {ltm_info['ann_label']}")

    # Step 4: Stock prices & market cap
    if not skip_prices:
        fy_end_dates = get_fiscal_year_end_dates(facts, years)
        closing_prices = get_historical_closing_prices(company_info['ticker'], fy_end_dates)
    else:
        closing_prices = {yr: None for yr in years}

    market_cap_data = {}
    shares = financial_data['income_statement']['shares_diluted']
    for yr in years:
        price = closing_prices.get(yr)
        shr   = shares.get(yr)
        if price is not None and shr is not None:
            market_cap_data[yr] = price * shr
        else:
            market_cap_data[yr] = None

    financial_data['market_cap'] = market_cap_data
    financial_data['stock_prices'] = closing_prices

    # Step 4b: WACC data — share price date, treasury yield, ERP, comparable companies
    if not skip_prices:
        treasury_yield = get_treasury_yield()
        kroll_erp = get_kroll_erp()
    else:
        treasury_yield = None
        kroll_erp = None

    rf = treasury_yield or 0.045
    erp = kroll_erp or 0.05

    comp_data = []
    if not auto_select and not skip_prices and price_date is None:
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
    bs  = financial_data['balance_sheet']
    basic_sh   = (inc['shares_basic'].get(latest) or 0) / 1e6
    diluted_sh = (inc['shares_diluted'].get(latest) or 0) / 1e6
    int_exp    = abs(inc['interest_expense'].get(latest) or 0)
    total_debt_raw = (bs['st_debt'].get(latest) or 0) + (bs['lt_debt'].get(latest) or 0)
    implied_cod = round(int_exp / total_debt_raw, 4) if total_debt_raw > 0 else 0.05

    financial_data['wacc_inputs'] = {
        'current_price':         current_price_data,
        'treasury_yield':        rf,
        'kroll_erp':             erp,
        'comparables':           comp_data,
        'shares_breakdown': {
            'basic':     round(basic_sh, 2),
            'rsus':      round(max(diluted_sh - basic_sh, 0), 2),
            'options':   0,
            'conv_debt': 0,
            'conv_pref': 0,
        },
        'implied_cod': implied_cod,
    }

    # Step 5: Build Excel workbook
    safe = _safe_name(company_info['name'])
    output_dir = os.path.join(os.path.dirname(os.path.abspath(__file__)), 'models')
    os.makedirs(output_dir, exist_ok=True)
    output_file = os.path.join(output_dir, f"{safe}_Financial_Model.xlsx")

    # Save JSON backup
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


def bulk_test():
    """Run the model through the 20 largest companies and print a validation report."""
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
            result = build_model(ticker, skip_prices=True, auto_select=True)
            if result is None:
                results.append({'ticker': ticker, 'status': 'SKIP', 'reason': 'not found / no data'})
            else:
                company_info, val = result
                results.append({
                    'ticker':  ticker,
                    'name':    company_info['name'],
                    'status':  'OK',
                    'total':   val['total'],
                    'passed':  val['passed'],
                    'failed':  val['failed'],
                    'checks':  val['checks'],
                })
        except Exception as e:
            results.append({'ticker': ticker, 'status': 'ERROR', 'reason': str(e)})
            print(f"  [ERROR] {ticker}: {e}")

        # Respect SEC rate limit (10 req/sec, but be conservative)
        time.sleep(0.2)

    # -- Consolidated Report ------------------------------------------
    print()
    print()
    print("=" * 70)
    print("  BULK VALIDATION REPORT")
    print("=" * 70)
    print()

    grand_total  = 0
    grand_passed = 0
    grand_failed = 0

    for r in results:
        if r['status'] == 'OK':
            grand_total  += r['total']
            grand_passed += r['passed']
            grand_failed += r['failed']

            if r['failed'] == 0:
                print(f"  {r['ticker']:>6s}:  {r['passed']}/{r['total']} checks passed  "
                      f"- {r['name']}")
            else:
                # List failed checks
                fail_details = []
                years = r['checks'][0]['expected'].keys()
                for chk in r['checks']:
                    for yr in years:
                        exp = chk['expected'].get(yr)
                        der = chk['derived'].get(yr)
                        tol_ref_d = chk.get('tolerance_ref')
                        tol_r = tol_ref_d.get(yr, exp) if tol_ref_d else exp
                        passed, diff = _check_status(exp, der, tol_r,
                                                     tolerance_pct=chk.get('tolerance_pct'))
                        if not passed:
                            fail_details.append(f"{chk['name']} (FY{yr}, {diff:+.1f})")
                print(f"  {r['ticker']:>6s}:  {r['passed']}/{r['total']} checks passed  "
                      f"- {r['name']}")
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


def main():
    print()
    print("=" * 60)
    print("  Financial Model Builder  -  SEC EDGAR Edition")
    print("=" * 60)
    print()

    # Check for --bulk-test flag
    if '--bulk-test' in sys.argv:
        bulk_test()
        return

    # Allow company name/ticker as command-line argument
    args = [a for a in sys.argv[1:] if not a.startswith('--')]
    if args:
        company_name = ' '.join(args).strip()
        print(f"  Company: {company_name}")
    else:
        company_name = input("  Enter company name or ticker symbol: ").strip()

    if not company_name:
        print("  No company provided. Exiting.")
        sys.exit(1)

    # Step 1: Find company CIK
    print()
    print("[1/6] Searching SEC EDGAR...")
    company_info = search_company(company_name)
    if not company_info:
        print("\n  Could not find the company. Try the exact ticker (e.g., AAPL) or")
        print("  the official name as it appears in SEC filings.")
        sys.exit(1)

    print(f"\n  [OK] Company : {company_info['name']}")
    print(f"       Ticker  : {company_info['ticker']}")
    print(f"       CIK     : {company_info['cik']}")

    # Step 2: Download XBRL facts
    print()
    print("[2/6] Downloading XBRL financial data from SEC EDGAR...")
    facts = get_company_facts(company_info['cik'])
    if not facts:
        print("\n  Could not fetch financial data. The company may not have")
        print("  XBRL-tagged filings, or the SEC API may be temporarily unavailable.")
        sys.exit(1)

    # Determine the 4 most recent fiscal years
    years = get_fiscal_years(facts, n_years=4)
    if not years:
        print("\n  Could not find annual filing data for this company.")
        sys.exit(1)

    years_display = ', '.join(f'FY{y}' for y in reversed(years))
    print(f"  [OK] Found data for: {years_display}")

    # Step 3: Extract financial data
    print()
    print("[3/6] Extracting 3-statement financial data...")
    financial_data = extract_financial_data(facts, years, ticker=company_info['ticker'])

    # Quick sanity summary
    inc = financial_data['income_statement']
    latest_yr = years[0]
    rev = inc['revenue'].get(latest_yr)
    ni  = inc['net_income'].get(latest_yr)
    print(f"  [OK] Latest year FY{latest_yr}:")
    if rev:
        print(f"         Revenue    : ${rev/1e6:,.1f}M")
    if ni:
        print(f"         Net Income : ${ni/1e6:,.1f}M")

    # Step 4: Fetch stock prices & compute market cap
    print()
    print("[4/6] Fetching historical stock prices (Yahoo Finance)...")
    fy_end_dates = get_fiscal_year_end_dates(facts, years)
    closing_prices = get_historical_closing_prices(company_info['ticker'], fy_end_dates)

    market_cap_data = {}
    shares = financial_data['income_statement']['shares_diluted']
    for yr in years:
        price = closing_prices.get(yr)
        shr   = shares.get(yr)
        if price is not None and shr is not None:
            market_cap_data[yr] = price * shr
            print(f"  [OK] FY{yr} Market Cap: ${market_cap_data[yr]/1e6:,.0f}M "
                  f"(${price:.2f} x {shr/1e6:,.1f}M shares)")
        else:
            market_cap_data[yr] = None
            print(f"  [--] FY{yr} Market Cap: unavailable (will use book equity)")

    financial_data['market_cap'] = market_cap_data
    financial_data['stock_prices'] = closing_prices

    # Step 5: WACC data — share price date, treasury yield, ERP, comparable companies
    print()
    print("[5/6] Fetching WACC inputs...")
    treasury_yield = get_treasury_yield()
    kroll_erp = get_kroll_erp()

    rf_main = treasury_yield or 0.045
    erp_main = kroll_erp or 0.05

    if treasury_yield:
        print(f"  [OK] Risk-Free Rate (CNBC US10Y): {treasury_yield*100:.2f}%")
    if kroll_erp:
        print(f"  [OK] Equity Risk Premium (Kroll): {kroll_erp*100:.1f}%")

    print()
    price_date = input("  Share price date for WACC (YYYY-MM-DD, or Enter for latest): ").strip() or None

    print("  Finding comparable companies...")
    comp_data, industry_name = get_industry_peers(company_info['ticker'])
    if comp_data:
        print(f"  [OK] Comparable companies ({industry_name}):")
        for cd in comp_data:
            print(f"    {cd['name']} ({cd['ticker']}): beta={cd['beta']:.2f}, ${cd['price']:.2f}")
    else:
        print("  [--] No comparable companies found")

    current_price_data = get_current_price(company_info['ticker'], price_date)
    if current_price_data['price']:
        print(f"  [OK] Share Price: ${current_price_data['price']:.2f} ({current_price_data['date']})")
    else:
        print("  [--] Share price unavailable")

    inc_w = financial_data['income_statement']
    bs_w  = financial_data['balance_sheet']
    latest_w = years[0]
    basic_sh   = (inc_w['shares_basic'].get(latest_w) or 0) / 1e6
    diluted_sh = (inc_w['shares_diluted'].get(latest_w) or 0) / 1e6
    int_exp_w  = abs(inc_w['interest_expense'].get(latest_w) or 0)
    total_debt_w = (bs_w['st_debt'].get(latest_w) or 0) + (bs_w['lt_debt'].get(latest_w) or 0)
    implied_cod = round(int_exp_w / total_debt_w, 4) if total_debt_w > 0 else 0.05

    financial_data['wacc_inputs'] = {
        'current_price':         current_price_data,
        'treasury_yield':        rf_main,
        'kroll_erp':             erp_main,
        'comparables':           comp_data,
        'shares_breakdown': {
            'basic':     round(basic_sh, 2),
            'rsus':      round(max(diluted_sh - basic_sh, 0), 2),
            'options':   0,
            'conv_debt': 0,
            'conv_pref': 0,
        },
        'implied_cod': implied_cod,
    }

    # Step 6: Build Excel workbook
    print()
    print("[6/6] Building Excel financial model...")

    safe_name = _safe_name(company_info['name'])
    output_dir = os.path.join(os.path.dirname(os.path.abspath(__file__)), 'models')
    os.makedirs(output_dir, exist_ok=True)
    output_file = os.path.join(output_dir, f"{safe_name}_Financial_Model.xlsx")

    # JSON requires string keys; fiscal years are ints in the data dicts.
    def _stringify_keys(obj):
        if isinstance(obj, dict):
            return {str(k): _stringify_keys(v) for k, v in obj.items()}
        return obj

    tenk_dir = os.path.join(output_dir, '10k')
    os.makedirs(tenk_dir, exist_ok=True)

    # Save extracted financial data as JSON.
    tenk_file = os.path.join(tenk_dir, f"{safe_name}_10k.json")
    with open(tenk_file, 'w', encoding='utf-8') as f:
        json.dump({
            'company': company_info,
            'fiscal_years': years,
            'financial_data': _stringify_keys(financial_data),
        }, f, indent=2)
    print(f"  [OK] 10-K data saved to: models/10k/{safe_name}_10k.json")

    validation_results = create_excel(company_info, financial_data, output_file)

    abs_path = os.path.abspath(output_file)
    print()
    print("=" * 60)
    print("  Done!")
    print(f"  File saved to: {abs_path}")
    print("=" * 60)
    print()
    print("  Excel workbook contents:")
    print("  - Sheet 'Financial Statements': 3-statement model (historical)")
    print("  - Sheet 'WACC'               : Weighted avg cost of capital")
    print("  - Sheet 'DCF Model'           : 5-year DCF valuation")
    print("  - Sheet 'Data Validation'     : 12 cross-checks (PASS/FAIL)")
    print()
    if validation_results:
        p = validation_results['passed']
        t = validation_results['total']
        f = validation_results['failed']
        print(f"  Validation: {p}/{t} checks passed", end='')
        if f > 0:
            print(f"  ({f} failed - see Data Validation sheet)")
        else:
            print("  (all passed)")
    print()
    print("  Tip: Yellow cells in the WACC and DCF sheets are editable inputs.")
    print("       Change WACC assumptions and growth rates to run scenarios.")
    print()


if __name__ == '__main__':
    main()
