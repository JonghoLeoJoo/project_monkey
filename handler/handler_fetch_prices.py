"""Step 4 – 과거 주가 조회 및 시가총액 계산."""

from typing import Dict, List

from sec_fetcher import get_fiscal_year_end_dates, get_historical_closing_prices


def run(facts: Dict, years: List[int], company_info: Dict, financial_data: Dict) -> Dict:
    """각 회계연도 종가를 조회하고 시가총액을 계산해 financial_data에 추가한다.

    Returns:
        financial_data (in-place 수정 후 반환)
    """
    print()
    print("[4/6] Fetching historical stock prices (Yahoo Finance)...")
    fy_end_dates    = get_fiscal_year_end_dates(facts, years)
    closing_prices  = get_historical_closing_prices(company_info['ticker'], fy_end_dates)

    shares = financial_data['income_statement']['shares_diluted']
    market_cap_data = {}
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

    financial_data['market_cap']    = market_cap_data
    financial_data['stock_prices']  = closing_prices
    return financial_data
