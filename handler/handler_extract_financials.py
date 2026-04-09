"""Step 3 – 재무제표 데이터 추출."""

from typing import Dict, List

from sec_fetcher import extract_financial_data


def run(facts: Dict, years: List[int], ticker: str) -> Dict:
    """XBRL facts에서 3-statement 재무 데이터를 추출한다.

    Returns:
        financial_data dict
    """
    print()
    print("[3/6] Extracting 3-statement financial data...")
    financial_data = extract_financial_data(facts, years, ticker=ticker)

    inc = financial_data['income_statement']
    latest_yr = years[0]
    rev = inc['revenue'].get(latest_yr)
    ni  = inc['net_income'].get(latest_yr)
    print(f"  [OK] Latest year FY{latest_yr}:")
    if rev:
        print(f"         Revenue    : ${rev/1e6:,.1f}M")
    if ni:
        print(f"         Net Income : ${ni/1e6:,.1f}M")

    return financial_data
