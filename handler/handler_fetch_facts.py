"""Step 2 – SEC EDGAR XBRL 데이터 다운로드."""

import sys
from typing import Dict, List, Tuple

from sec_fetcher import get_company_facts, get_fiscal_years


def run(company_info: Dict[str, str]) -> Tuple[Dict, List[int]]:
    """회사 CIK로 XBRL facts를 가져오고 최근 4개 회계연도를 추출한다.

    Returns:
        (facts, years)  – years: 최신 순 정렬된 회계연도 리스트
    Raises:
        SystemExit: 데이터 수신 실패 또는 연도 없음
    """
    print()
    print("[2/6] Downloading XBRL financial data from SEC EDGAR...")
    facts = get_company_facts(company_info['cik'])
    if not facts:
        print("\n  Could not fetch financial data. The company may not have\n"
              "  XBRL-tagged filings, or the SEC API may be temporarily unavailable.")
        sys.exit(1)

    years = get_fiscal_years(facts, n_years=4)
    if not years:
        print("\n  Could not find annual filing data for this company.")
        sys.exit(1)

    years_display = ', '.join(f'FY{y}' for y in reversed(years))
    print(f"  [OK] Found data for: {years_display}")
    return facts, years
