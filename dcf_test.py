"""
DCF Accuracy Test
=================
Runs the financial model extraction on the 50 largest US public companies,
computes the DCF-implied share price in Python, and compares it against
the actual share price as of 2025-12-19.

Companies with wildly inaccurate implied prices (>2x or <0.5x actual, or
negative) are flagged for investigation of missing financial items.

Usage:
    python dcf_test.py
"""

import sys
import os
import json
import time
import traceback
from typing import Dict, Optional, List, Tuple

from sec_fetcher import (
    search_company,
    get_company_facts,
    get_fiscal_years,
    extract_financial_data,
    get_fiscal_year_end_dates,
    get_historical_closing_prices,
    get_current_price,
    get_treasury_yield,
    get_kroll_erp,
    get_industry_peers,
)

# 50 largest US public companies by market cap (as of late 2025)
TEST_TICKERS = [
    'AAPL', 'MSFT', 'NVDA', 'GOOG', 'AMZN',
    'META', 'BRK-B', 'LLY', 'AVGO', 'JPM',
    'TSLA', 'WMT', 'V', 'UNH', 'XOM',
    'MA', 'COST', 'PG', 'JNJ', 'HD',
    'ORCL', 'ABBV', 'MRK', 'CRM', 'BAC',
    'NFLX', 'AMD', 'CVX', 'KO', 'TMO',
    'LIN', 'PEP', 'ADBE', 'CSCO', 'WFC',
    'ACN', 'ABT', 'MCD', 'GE', 'DHR',
    'NOW', 'TXN', 'IBM', 'PM', 'INTU',
    'AMGN', 'QCOM', 'CAT', 'DIS', 'NEE',
]

PRICE_DATE = '2025-12-19'


def compute_dcf_price(fd: Dict, comp_data: List[Dict],
                      rf: float, erp: float,
                      current_price: float) -> Dict:
    """Replicate the Excel DCF calculation in Python.

    Returns a dict with intermediate values and the implied share price.
    """
    years_desc = fd['years']
    latest_yr = years_desc[0]
    inc = fd['income_statement']
    bs = fd['balance_sheet']
    cf = fd['cash_flow']

    def v(d, yr):
        val = d.get(yr)
        return val / 1e6 if val is not None else None

    # --- Projection assumptions (same as _calc_dcf_assumptions) ---
    # Revenue growth
    rev_growths = []
    for i in range(len(years_desc) - 1):
        r0 = v(inc['revenue'], years_desc[i + 1])
        r1 = v(inc['revenue'], years_desc[i])
        if r0 and r1 and r0 > 0:
            rev_growths.append((r1 / r0) - 1.0)
    avg_rev_growth = sum(rev_growths) / len(rev_growths) if rev_growths else 0.05

    # EBITDA margin
    ebitda_margins = []
    for yr in years_desc:
        rev = v(inc['revenue'], yr)
        ebt = v(inc['ebitda'], yr)
        if rev and ebt:
            ebitda_margins.append(ebt / rev)
    avg_ebitda_margin = sum(ebitda_margins) / len(ebitda_margins) if ebitda_margins else 0.20

    # D&A %
    da_pcts = []
    for yr in years_desc:
        rev = v(inc['revenue'], yr)
        da = v(inc['da'], yr)
        if rev and da:
            da_pcts.append(da / rev)
    avg_da_pct = sum(da_pcts) / len(da_pcts) if da_pcts else 0.05

    # Capex %
    capex_pcts = []
    for yr in years_desc:
        rev = v(inc['revenue'], yr)
        capx = v(cf['capex'], yr)
        if rev and capx:
            capex_pcts.append(abs(capx) / rev)
    avg_capex_pct = sum(capex_pcts) / len(capex_pcts) if capex_pcts else 0.04

    # Tax rate
    tax_rates = []
    for yr in years_desc:
        pre = v(inc['pretax_income'], yr)
        tax = v(inc['tax_expense'], yr)
        if pre and tax and pre > 0:
            tax_rates.append(tax / pre)
    avg_tax = sum(tax_rates) / len(tax_rates) if tax_rates else 0.21

    # --- WACC calculation ---
    # Cost of debt
    int_exp_raw = abs(inc['interest_expense'].get(latest_yr) or 0)
    total_debt_raw = (bs['st_debt'].get(latest_yr) or 0) + (bs['lt_debt'].get(latest_yr) or 0)
    cost_of_debt = (int_exp_raw / total_debt_raw) if total_debt_raw > 0 else 0.05
    cod_after_tax = cost_of_debt * (1 - avg_tax)

    # Beta from comparable companies (un-lever, average, re-lever)
    delevered_betas = []
    for comp in comp_data:
        beta = comp.get('beta')
        mkt_cap = comp.get('price', 0) * comp.get('shares', 0)
        net_debt = comp.get('net_debt', 0)
        tax_rate = comp.get('tax_rate', 0.21)
        if beta and mkt_cap > 0:
            denom = 1 + (1 - tax_rate) * net_debt / mkt_cap
            if denom > 0:
                delevered_betas.append(beta / denom)

    if len(delevered_betas) > 2:
        # Trim high/low outliers
        sorted_betas = sorted(delevered_betas)
        trimmed = sorted_betas[1:-1]
        avg_delev_beta = sum(trimmed) / len(trimmed)
    elif delevered_betas:
        avg_delev_beta = sum(delevered_betas) / len(delevered_betas)
    else:
        avg_delev_beta = 1.0

    # Re-lever for target company
    shares_diluted = (inc['shares_diluted'].get(latest_yr) or 0) / 1e6
    if current_price and shares_diluted > 0:
        tgt_mktcap = current_price * shares_diluted
    else:
        tgt_mktcap = (v(bs['total_equity'], latest_yr) or 0)

    st_debt_m = (v(bs['st_debt'], latest_yr) or 0)
    lt_debt_m = (v(bs['lt_debt'], latest_yr) or 0)
    cash_m = (v(bs['cash'], latest_yr) or 0)
    net_debt_m = st_debt_m + lt_debt_m - cash_m

    if tgt_mktcap > 0:
        relevered_beta = avg_delev_beta * (1 + (1 - avg_tax) * net_debt_m / tgt_mktcap)
    else:
        relevered_beta = avg_delev_beta

    cost_of_equity = rf + relevered_beta * erp

    # Capital structure weights
    ev = tgt_mktcap + net_debt_m
    if ev > 0:
        eq_weight = tgt_mktcap / ev
        debt_weight = net_debt_m / ev
    else:
        eq_weight = 1.0
        debt_weight = 0.0

    wacc = eq_weight * cost_of_equity + debt_weight * cod_after_tax

    # Sanity: if WACC is too low or negative, cap it
    if wacc <= 0.02:
        wacc = 0.10  # fallback

    terminal_growth = 0.025

    # --- 5-Year Projections ---
    base_revenue = v(inc['revenue'], latest_yr) or 0

    revenues = []
    ebitdas = []
    das = []
    ebits = []
    nopats = []
    capexes = []
    ufcfs = []
    pv_fcfs = []

    rev = base_revenue
    for i in range(5):
        rev = rev * (1 + avg_rev_growth)
        revenues.append(rev)
        ebitda = rev * avg_ebitda_margin
        ebitdas.append(ebitda)
        da = rev * avg_da_pct
        das.append(da)
        ebit = ebitda - da
        ebits.append(ebit)
        nopat = ebit * (1 - avg_tax)
        nopats.append(nopat)
        capex = rev * avg_capex_pct
        capexes.append(capex)
        ufcf = nopat + da - capex
        ufcfs.append(ufcf)
        disc = 1 / (1 + wacc) ** (i + 1)
        pv_fcfs.append(ufcf * disc)

    sum_pv_fcf = sum(pv_fcfs)

    # Terminal value
    if wacc > terminal_growth:
        terminal_value = ufcfs[-1] * (1 + terminal_growth) / (wacc - terminal_growth)
    else:
        terminal_value = ufcfs[-1] * 20  # fallback multiplier

    pv_tv = terminal_value / (1 + wacc) ** 5
    enterprise_value = sum_pv_fcf + pv_tv
    equity_value = enterprise_value - net_debt_m

    if shares_diluted > 0:
        implied_price = equity_value / shares_diluted
    else:
        implied_price = None

    return {
        'implied_price': implied_price,
        'equity_value': equity_value,
        'enterprise_value': enterprise_value,
        'net_debt': net_debt_m,
        'shares_diluted': shares_diluted,
        'wacc': wacc,
        'cost_of_equity': cost_of_equity,
        'cost_of_debt': cost_of_debt,
        'relevered_beta': relevered_beta,
        'avg_delev_beta': avg_delev_beta,
        'terminal_growth': terminal_growth,
        'avg_rev_growth': avg_rev_growth,
        'avg_ebitda_margin': avg_ebitda_margin,
        'avg_da_pct': avg_da_pct,
        'avg_capex_pct': avg_capex_pct,
        'avg_tax': avg_tax,
        'base_revenue': base_revenue,
        'latest_yr': latest_yr,
        # Diagnostics - key financial items
        'revenue': v(inc['revenue'], latest_yr),
        'cogs': v(inc['cogs'], latest_yr),
        'operating_income': v(inc['operating_income'], latest_yr),
        'ebitda_val': v(inc['ebitda'], latest_yr),
        'da_val': v(inc['da'], latest_yr),
        'net_income': v(inc['net_income'], latest_yr),
        'operating_cf': v(cf['operating_cf'], latest_yr),
        'capex_val': v(cf['capex'], latest_yr),
        'fcf_val': v(cf['fcf'], latest_yr),
        'total_assets': v(bs['total_assets'], latest_yr),
        'total_liabilities': v(bs['total_liabilities'], latest_yr),
        'total_equity': v(bs['total_equity'], latest_yr),
        'cash_val': cash_m,
        'st_debt_val': st_debt_m,
        'lt_debt_val': lt_debt_m,
    }


def diagnose_missing_items(fd: Dict) -> List[str]:
    """Check for common missing financial items that could cause DCF inaccuracy."""
    years = fd['years']
    latest = years[0]
    inc = fd['income_statement']
    bs = fd['balance_sheet']
    cf = fd['cash_flow']

    issues = []

    def check(name, data, yr=latest):
        val = data.get(yr)
        if val is None or val == 0:
            issues.append(f"MISSING: {name} (FY{yr})")
            return False
        return True

    # Critical income statement items
    check('Revenue', inc['revenue'])
    check('COGS', inc['cogs'])
    check('Operating Income', inc['operating_income'])
    check('D&A', inc['da'])
    check('Pre-tax Income', inc['pretax_income'])
    check('Tax Expense', inc['tax_expense'])
    check('Net Income', inc['net_income'])
    check('Shares Diluted', inc['shares_diluted'])
    check('Interest Expense', inc['interest_expense'])

    # Critical balance sheet items
    check('Cash', bs['cash'])
    check('Total Assets', bs['total_assets'])
    check('Total Equity', bs['total_equity'])
    check('Total Current Assets', bs['total_current_a'])
    check('Total Current Liabilities', bs['total_current_l'])
    check('PPE Net', bs['ppe_net'])
    check('Accounts Receivable', bs['accounts_rec'])
    check('Accounts Payable', bs['accounts_pay'])

    # Critical cash flow items
    check('Operating CF', cf['operating_cf'])
    check('Capex', cf['capex'])
    check('Investing CF', cf['investing_cf'])
    check('Financing CF', cf['financing_cf'])

    # Check EBITDA derivation
    ebitda = inc['ebitda'].get(latest)
    if ebitda is None:
        issues.append("MISSING: EBITDA (could not derive from EBIT + D&A)")

    # Check for very large 'Other Operating Expenses' plug (sign of missing items)
    rev = inc['revenue'].get(latest) or 1
    cogs = inc['cogs'].get(latest) or 0
    rd = inc['rd_expense'].get(latest) or 0
    sga = inc['sga_expense'].get(latest) or 0
    ebit = inc['operating_income'].get(latest) or 0
    gp = rev - cogs
    other_opex = gp - rd - sga - ebit
    if rev > 0 and abs(other_opex) > 0.3 * rev:
        issues.append(f"LARGE PLUG: Other Operating Expenses = ${other_opex/1e6:,.0f}M "
                      f"({abs(other_opex)/rev*100:.0f}% of revenue) - suggests missing opex items")

    # Check BS balance
    ta = bs['total_assets'].get(latest) or 0
    te = bs['total_equity'].get(latest) or 0
    tl = bs['total_liabilities'].get(latest) or 0
    if ta > 0 and abs(ta - te - tl) > 0.01 * ta:
        issues.append(f"BS IMBALANCE: Assets={ta/1e6:,.0f}M vs L+E={((te+tl)/1e6):,.0f}M")

    # Check CF reconciliation
    opcf = cf['operating_cf'].get(latest) or 0
    invcf = cf['investing_cf'].get(latest) or 0
    fincf = cf['financing_cf'].get(latest) or 0
    fx = cf['fx_effect'].get(latest) or 0
    net_cf = opcf + invcf + fincf + fx
    if opcf != 0 and abs(net_cf) > 0.5 * abs(opcf):
        issues.append(f"CF RECONCILIATION: Net CF = ${net_cf/1e6:,.0f}M "
                      f"(vs OpCF=${opcf/1e6:,.0f}M)")

    return issues


def run_test():
    """Main test loop."""
    print()
    print("=" * 80)
    print("  DCF ACCURACY TEST - 50 Largest Public Companies")
    print(f"  Price Date: {PRICE_DATE}")
    print("=" * 80)
    print()

    # Fetch WACC inputs once (shared across all companies)
    treasury_yield = get_treasury_yield()
    kroll_erp = get_kroll_erp()
    rf = treasury_yield or 0.045
    erp = kroll_erp or 0.05
    print(f"  Risk-Free Rate: {rf*100:.2f}%")
    print(f"  Equity Risk Premium: {erp*100:.1f}%")
    print()

    results = []
    for i, ticker in enumerate(TEST_TICKERS, 1):
        print(f"\n{'─' * 80}")
        print(f"  [{i}/{len(TEST_TICKERS)}] {ticker}")
        print(f"{'─' * 80}")

        result = {
            'ticker': ticker,
            'status': None,
            'actual_price': None,
            'implied_price': None,
            'ratio': None,
            'issues': [],
            'error': None,
        }

        try:
            # 1. Find company
            company_info = search_company(ticker, auto_select=True)
            if not company_info:
                result['status'] = 'SKIP'
                result['error'] = 'Company not found'
                results.append(result)
                continue
            result['name'] = company_info['name']

            # 2. Get XBRL facts
            facts = get_company_facts(company_info['cik'])
            if not facts:
                result['status'] = 'SKIP'
                result['error'] = 'No XBRL data'
                results.append(result)
                continue

            # 3. Get fiscal years
            years = get_fiscal_years(facts, n_years=4)
            if not years:
                result['status'] = 'SKIP'
                result['error'] = 'No fiscal years found'
                results.append(result)
                continue

            # 4. Extract financial data
            financial_data = extract_financial_data(facts, years,
                                                    ticker=company_info['ticker'])

            # 5. Get actual price at 2025-12-19
            price_data = get_current_price(company_info['ticker'], PRICE_DATE)
            actual_price = price_data.get('price')
            result['actual_price'] = actual_price

            if not actual_price:
                result['status'] = 'SKIP'
                result['error'] = 'No price data for 2025-12-19'
                results.append(result)
                continue

            # 6. Get comparable companies for WACC
            comp_data, industry_name = get_industry_peers(company_info['ticker'])

            # 7. Compute DCF implied price
            dcf = compute_dcf_price(financial_data, comp_data, rf, erp, actual_price)
            implied_price = dcf.get('implied_price')
            result['implied_price'] = implied_price
            result['dcf'] = dcf

            if implied_price is None or implied_price <= 0:
                result['status'] = 'OUTLIER'
                result['ratio'] = None
                result['issues'] = diagnose_missing_items(financial_data)
            else:
                ratio = implied_price / actual_price
                result['ratio'] = ratio
                if ratio > 2.0 or ratio < 0.5:
                    result['status'] = 'OUTLIER'
                    result['issues'] = diagnose_missing_items(financial_data)
                else:
                    result['status'] = 'OK'

            # Print summary
            print(f"  {company_info['name']} ({ticker})")
            print(f"    Actual Price (2025-12-19): ${actual_price:.2f}")
            if implied_price is not None:
                print(f"    DCF Implied Price:         ${implied_price:.2f}")
                if result['ratio']:
                    print(f"    Ratio (Implied/Actual):    {result['ratio']:.2f}x")
            else:
                print(f"    DCF Implied Price:         N/A")
            if result['status'] == 'OUTLIER':
                print(f"    *** OUTLIER ***")
                for issue in result['issues']:
                    print(f"      - {issue}")

            # Print key DCF inputs for debugging
            print(f"    WACC: {dcf['wacc']*100:.1f}%, "
                  f"Rev Growth: {dcf['avg_rev_growth']*100:.1f}%, "
                  f"EBITDA Margin: {dcf['avg_ebitda_margin']*100:.1f}%")
            ebitda_s = f"${dcf['ebitda_val']:,.0f}M" if dcf['ebitda_val'] else 'N/A'
            ni_s = f"${dcf['net_income']:,.0f}M" if dcf['net_income'] else 'N/A'
            rev_s = f"${dcf['revenue']:,.0f}M" if dcf['revenue'] else 'N/A'
            print(f"    Revenue: {rev_s}, EBITDA: {ebitda_s}, Net Income: {ni_s}")
            opcf_s = f"${dcf['operating_cf']:,.0f}M" if dcf['operating_cf'] else 'N/A'
            capex_s = f"${dcf['capex_val']:,.0f}M" if dcf['capex_val'] else 'N/A'
            fcf_s = f"${dcf['fcf_val']:,.0f}M" if dcf['fcf_val'] else 'N/A'
            print(f"    OpCF: {opcf_s}, Capex: {capex_s}, FCF: {fcf_s}")

        except Exception as e:
            result['status'] = 'ERROR'
            result['error'] = str(e)
            print(f"  [ERROR] {ticker}: {e}")
            traceback.print_exc()

        results.append(result)
        time.sleep(0.3)  # SEC rate limit

    # ── Summary Report ─────────────────────────────────────────────────────
    print()
    print()
    print("=" * 80)
    print("  DCF ACCURACY TEST - SUMMARY REPORT")
    print("=" * 80)
    print()
    print(f"  {'Ticker':<8s} {'Name':<30s} {'Actual':>10s} {'Implied':>10s} {'Ratio':>8s} {'Status':>10s}")
    print(f"  {'─'*8} {'─'*30} {'─'*10} {'─'*10} {'─'*8} {'─'*10}")

    ok_count = 0
    outlier_count = 0
    skip_count = 0
    error_count = 0

    for r in results:
        name = r.get('name', r['ticker'])[:30]
        actual = f"${r['actual_price']:.2f}" if r['actual_price'] else 'N/A'
        implied = f"${r['implied_price']:.2f}" if r.get('implied_price') else 'N/A'
        ratio = f"{r['ratio']:.2f}x" if r.get('ratio') else 'N/A'
        status = r['status']

        flag = ''
        if status == 'OUTLIER':
            flag = ' ***'
            outlier_count += 1
        elif status == 'OK':
            ok_count += 1
        elif status == 'SKIP':
            skip_count += 1
        elif status == 'ERROR':
            error_count += 1

        print(f"  {r['ticker']:<8s} {name:<30s} {actual:>10s} {implied:>10s} {ratio:>8s} {status:>10s}{flag}")

    print()
    print(f"  OK: {ok_count}  |  OUTLIER: {outlier_count}  |  SKIP: {skip_count}  |  ERROR: {error_count}")
    print()

    # ── Outlier Detail ──────────────────────────────────────────────────────
    outliers = [r for r in results if r['status'] == 'OUTLIER']
    if outliers:
        print()
        print("=" * 80)
        print("  OUTLIER ANALYSIS - Companies with >2x or <0.5x ratio")
        print("=" * 80)
        for r in outliers:
            print(f"\n  {r['ticker']} - {r.get('name', '???')}")
            actual_s = f"${r['actual_price']:.2f}" if r['actual_price'] else 'N/A'
            implied_s = f"${r['implied_price']:.2f}" if r.get('implied_price') else 'N/A'
            print(f"    Actual: {actual_s}")
            print(f"    Implied: {implied_s}")
            if r.get('ratio'):
                print(f"    Ratio: {r['ratio']:.2f}x")
            if r.get('dcf'):
                dcf = r['dcf']
                print(f"    WACC: {dcf['wacc']*100:.1f}%")
                print(f"    Revenue Growth: {dcf['avg_rev_growth']*100:.1f}%")
                print(f"    EBITDA Margin: {dcf['avg_ebitda_margin']*100:.1f}%")
                print(f"    D&A %: {dcf['avg_da_pct']*100:.1f}%")
                print(f"    Capex %: {dcf['avg_capex_pct']*100:.1f}%")
                print(f"    Tax Rate: {dcf['avg_tax']*100:.1f}%")
                print(f"    Net Debt: ${dcf['net_debt']:,.0f}M")
                print(f"    Shares (Diluted): {dcf['shares_diluted']:,.1f}M")
            if r.get('issues'):
                print(f"    Issues:")
                for issue in r['issues']:
                    print(f"      - {issue}")

    # Save results to JSON
    output_file = os.path.join(os.path.dirname(os.path.abspath(__file__)),
                               'dcf_test_results.json')
    # Serialize safely
    safe_results = []
    for r in results:
        sr = {k: v for k, v in r.items() if k != 'dcf'}
        if r.get('dcf'):
            sr['dcf_summary'] = {
                'wacc': r['dcf']['wacc'],
                'avg_rev_growth': r['dcf']['avg_rev_growth'],
                'avg_ebitda_margin': r['dcf']['avg_ebitda_margin'],
                'avg_da_pct': r['dcf']['avg_da_pct'],
                'avg_capex_pct': r['dcf']['avg_capex_pct'],
                'avg_tax': r['dcf']['avg_tax'],
                'net_debt': r['dcf']['net_debt'],
                'revenue': r['dcf']['revenue'],
                'ebitda_val': r['dcf']['ebitda_val'],
                'net_income': r['dcf']['net_income'],
                'operating_cf': r['dcf']['operating_cf'],
                'capex_val': r['dcf']['capex_val'],
                'total_assets': r['dcf']['total_assets'],
                'total_equity': r['dcf']['total_equity'],
                'base_revenue': r['dcf']['base_revenue'],
                'shares_diluted': r['dcf']['shares_diluted'],
                'enterprise_value': r['dcf']['enterprise_value'],
                'equity_value': r['dcf']['equity_value'],
            }
        safe_results.append(sr)

    with open(output_file, 'w') as f:
        json.dump(safe_results, f, indent=2)
    print(f"\n  Results saved to: {output_file}")


if __name__ == '__main__':
    run_test()
