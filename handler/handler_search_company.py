"""Step 1 – SEC EDGAR 회사 검색."""

import sys
import logging
from typing import Dict

from sec_fetcher import search_company

LOGGER = logging.getLogger(__name__)


def run(company_name: str) -> Dict[str, str]:
    """회사명 또는 티커로 SEC EDGAR를 검색한다.

    Returns:
        {'cik': str, 'name': str, 'ticker': str}
    Raises:
        SystemExit: 검색 실패 시
    """
    print()
    print("[1/6] Searching SEC EDGAR...")
    try:
        company_info = search_company(company_name)
    except Exception as e:
        LOGGER.error("%s", e)
        sys.exit(1)

    print(f"\n  [OK] Company : {company_info['name']}")
    print(f"       Ticker  : {company_info['ticker']}")
    print(f"       CIK     : {company_info['cik']}")
    return company_info
