"""Step 5 – WACC 입력값 수집 (금리, ERP, 비교 기업, 주가)."""

import json
from pathlib import Path
from typing import Dict, List, Optional

from sec_fetcher import (
    get_treasury_yield,
    get_kroll_erp,
    get_industry_peers,
    get_current_price,
)

CONFIG_PATH = Path(__file__).parent.parent / 'config' / 'main_config.json'


def _load_defaults() -> Dict:
    with open(CONFIG_PATH, 'r', encoding='utf-8') as f:
        return json.load(f)['defaults']


def run(years: List[int], company_info: Dict, financial_data: Dict,
        price_date: Optional[str] = None) -> Dict:
    """Treasury yield, Kroll ERP, 비교 기업, 현재 주가를 수집해 financial_data에 추가한다.

    Args:
        price_date: 'YYYY-MM-DD' 또는 None (최신 종가 사용)

    Returns:
        financial_data (in-place 수정 후 반환)
    """
    defaults = _load_defaults()
    rf_default  = defaults['risk_free_rate']
    erp_default = defaults['equity_risk_premium']
    cod_default = defaults['implied_cost_of_debt']

    print()
    print("[5/6] Fetching WACC inputs...")

    treasury_yield = get_treasury_yield()
    kroll_erp      = get_kroll_erp()
    rf  = treasury_yield or rf_default
    erp = kroll_erp      or erp_default

    if treasury_yield:
        print(f"  [OK] Risk-Free Rate (CNBC US10Y): {treasury_yield*100:.2f}%")
    if kroll_erp:
        print(f"  [OK] Equity Risk Premium (Kroll): {kroll_erp*100:.1f}%")

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

    inc     = financial_data['income_statement']
    bs      = financial_data['balance_sheet']
    latest  = years[0]
    basic_sh   = (inc['shares_basic'].get(latest)   or 0) / 1e6
    diluted_sh = (inc['shares_diluted'].get(latest) or 0) / 1e6
    int_exp    = abs(inc['interest_expense'].get(latest) or 0)
    total_debt = (bs['st_debt'].get(latest) or 0) + (bs['lt_debt'].get(latest) or 0)
    implied_cod = round(int_exp / total_debt, 4) if total_debt > 0 else cod_default

    financial_data['wacc_inputs'] = {
        'current_price':    current_price_data,
        'treasury_yield':   rf,
        'kroll_erp':        erp,
        'comparables':      comp_data,
        'shares_breakdown': {
            'basic':     round(basic_sh, 2),
            'rsus':      round(max(diluted_sh - basic_sh, 0), 2),
            'options':   0,
            'conv_debt': 0,
            'conv_pref': 0,
        },
        'implied_cod': implied_cod,
    }
    return financial_data
