"""
Financial Model Builder
=======================
Fetches the latest 4 years of 10-K data from SEC EDGAR and generates
an Excel workbook with:
  - Sheet 1: Historical 3-Statement Model (Income Statement, Balance Sheet, Cash Flow)
  - Sheet 2: DCF Valuation (5-year projection + terminal value)
  - Sheet 3: Data Validation (12 cross-checks with PASS/FAIL status)

Usage:
    python main.py                          # 대화형 입력
    python main.py "Apple"                  # 전체 파이프라인 실행
    python main.py AAPL                     # 전체 파이프라인 실행 (+ 각 단계 캐시 저장)
    python main.py AAPL --from-step 4       # Step 4부터 재실행 (캐시에서 Step 3 결과 로드)
"""

import sys
import logging
import argparse

import handler.handler_search_company     as h1
import handler.handler_fetch_facts        as h2
import handler.handler_extract_financials as h3
import handler.handler_fetch_prices       as h4
import handler.handler_wacc_inputs        as h5
import handler.handler_build_excel        as h6
import handler.cache                      as cache

LOGGER = logging.getLogger(__name__)


def _parse_args():
    parser = argparse.ArgumentParser(
        description='Financial Model Builder – SEC EDGAR Edition',
        formatter_class=argparse.RawDescriptionHelpFormatter,
        epilog=(
            'examples:\n'
            '  python main.py AAPL                # 전체 실행 (각 단계 캐시 저장)\n'
            '  python main.py AAPL --from-step 4  # Step 3 캐시 로드 후 Step 4부터 재실행\n'
            '  python main.py AAPL --from-step 6  # Step 5 캐시 로드 후 Excel만 재생성\n'
        ),
    )
    parser.add_argument(
        'company', nargs='?',
        help='회사명 또는 티커 (예: AAPL, "Apple Inc.")',
    )
    parser.add_argument(
        '--from-step', type=int, choices=range(2, 7), metavar='N',
        help=(
            'Step N(2–6)부터 재실행. '
            '이전 단계 결과를 .cache/ 에서 로드한다. '
            'company 인자에 티커 심볼을 그대로 사용해야 캐시 파일을 찾을 수 있다.'
        ),
    )
    return parser.parse_args()


def main():
    args      = _parse_args()
    from_step = args.from_step or 1

    print(f"\n{'=' * 60}\n  Financial Model Builder  -  SEC EDGAR Edition\n{'=' * 60}\n")

    company_name = args.company
    if not company_name:
        company_name = input("  Enter company name or ticker symbol: ").strip()
    if not company_name:
        print("  No company provided. Exiting.")
        sys.exit(1)

    # --from-step 지정 시 이전 단계 결과를 캐시에서 복원
    company_info = facts = years = financial_data = None
    ticker = company_name.upper()

    if from_step > 1:
        try:
            state = cache.load(ticker, from_step - 1)
        except FileNotFoundError as e:
            LOGGER.error("%s", e)
            sys.exit(1)
        company_info   = state['company_info']
        facts          = state.get('facts')
        years          = state.get('years')
        financial_data = state.get('financial_data', {})
        ticker         = company_info['ticker']
        print(
            f"  Resuming from step {from_step} "
            f"(loaded step {from_step - 1} cache for {company_info['name']})\n"
        )

    # ── Step 1: Find company CIK ──────────────────────────────────────────
    if from_step <= 1:
        try:
            company_info = h1.run(company_name)
        except Exception as e:
            LOGGER.error("%s", e)
            sys.exit(1)
        ticker = company_info['ticker']
        cache.save(ticker, 1, {'company_info': company_info})

    # ── Step 2: Download XBRL facts ───────────────────────────────────────
    if from_step <= 2:
        facts, years = h2.run(company_info)
        if not facts or not years:
            sys.exit(1)
        cache.save(ticker, 2, {
            'company_info': company_info,
            'facts':        facts,
            'years':        years,
        })

    # ── Step 3: Extract financial data ────────────────────────────────────
    if from_step <= 3:
        financial_data = h3.run(facts, years, company_info['ticker'])
        cache.save(ticker, 3, {
            'company_info':   company_info,
            'facts':          facts,
            'years':          years,
            'financial_data': financial_data,
        })

    # ── Step 4: Fetch stock prices & compute market cap ───────────────────
    if from_step <= 4:
        financial_data = h4.run(facts, years, company_info, financial_data)
        cache.save(ticker, 4, {
            'company_info':   company_info,
            'facts':          facts,
            'years':          years,
            'financial_data': financial_data,
        })

    # ── Step 5: WACC inputs ───────────────────────────────────────────────
    if from_step <= 5:
        price_date = input(
            "\n  Share price date for WACC (YYYY-MM-DD, or Enter for latest): "
        ).strip() or None
        financial_data = h5.run(years, company_info, financial_data, price_date)
        cache.save(ticker, 5, {
            'company_info':   company_info,
            'facts':          facts,
            'years':          years,
            'financial_data': financial_data,
        })

    # ── Step 6: Build Excel workbook ──────────────────────────────────────
    h6.run(company_info, financial_data, years)


if __name__ == '__main__':
    logging.basicConfig(level=logging.INFO, format='%(levelname)s: %(message)s')
    try:
        main()
    except Exception:
        LOGGER.exception("Unhandled error in main")
        sys.exit(1)
