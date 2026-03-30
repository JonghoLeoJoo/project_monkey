"""
XBRL Tag Investigation
======================
For each problematic company, fetch XBRL data and find what tags they
actually use for the financial items our model is missing.
"""

import json
import sys
from pathlib import Path
from sec_fetcher import search_company, get_company_facts, get_fiscal_years, _get_annual_values

CONFIG_PATH = Path(__file__).parent / 'config' / 'xbrl_investigate_config.json'


def load_config(path=CONFIG_PATH):
    """Load investigation constants from external JSON config."""
    with open(path, 'r', encoding='utf-8') as f:
        return json.load(f)


CONFIG = load_config()
INVESTIGATE = CONFIG['investigate']
SEARCH_TAGS = CONFIG['search_tags']
SEARCH_KEYWORDS = CONFIG['search_keywords']
DEFAULT_PRIORITY = CONFIG['default_priority']
MILLION_FORMAT_THRESHOLD = CONFIG['formatting']['million_format_threshold']
SIGNIFICANT_VALUE_THRESHOLD = CONFIG['formatting']['significant_value_threshold']
SHARES_SIGNIFICANT_THRESHOLD = CONFIG['formatting']['shares_significant_threshold']
SLEEP_SECONDS = CONFIG.get('sleep_seconds', 0.3)


def scan_all_concepts(facts, years, category_tags):
    """Try all tags in a category and report which ones have data."""
    results = []

    for tag in category_tags:
        values = _get_annual_values(facts, tag)
        if values:
            year_data = {v['fy']: v['val'] for v in values if v['fy'] in years}
            if year_data:
                results.append((tag, year_data))

    return results


def scan_all_available_concepts(facts, years, keyword):
    """Scan ALL us-gaap concepts for a keyword match and report which have data."""
    results = []
    us_gaap = facts.get('facts', {}).get('us-gaap', {})

    for concept_name in us_gaap:
        if keyword.lower() in concept_name.lower():
            values = _get_annual_values(facts, concept_name)
            if values:
                year_data = {v['fy']: v['val'] for v in values if v['fy'] in years}
                if year_data:
                    results.append((concept_name, year_data))

    return results


def investigate_company(ticker, issues):
    """Investigate a single company's XBRL data."""
    print(f"\n{'='*80}")
    print(f"  {ticker}")
    print(f"{'='*80}")

    company_info = search_company(ticker, auto_select=True)
    if not company_info:
        print(f"  Could not find {ticker}")
        return

    facts = get_company_facts(company_info['cik'])
    if not facts:
        print(f"  Could not fetch XBRL data")
        return

    years = get_fiscal_years(facts, n_years=4)
    if not years:
        print(f"  No fiscal years found")
        return

    print(f"  Company: {company_info['name']}")
    print(f"  Fiscal years: {years}")

    for issue in issues:
        print(f"\n  --- Investigating: {issue} ---")

        # Check known tags
        if issue in SEARCH_TAGS:
            known_results = scan_all_concepts(facts, years, SEARCH_TAGS[issue])
            if known_results:
                print(f"  Known tags with data:")
                for tag, data in known_results:
                    latest = data.get(years[0])
                    val_str = (
                        f"${latest/1e6:,.0f}M"
                        if latest and abs(latest) > MILLION_FORMAT_THRESHOLD
                        else str(latest)
                    )
                    print(f"    {tag}: {val_str} (FY{years[0]})")
            else:
                print(f"  No known tags found data!")

        if issue in SEARCH_KEYWORDS:
            print(f"  Broad concept search:")
            for kw in SEARCH_KEYWORDS[issue]:
                broad_results = scan_all_available_concepts(facts, years, kw)
                for tag, data in broad_results:
                    latest = data.get(years[0])
                    if latest is None:
                        continue
                    val_str = (
                        f"${latest/1e6:,.1f}M"
                        if abs(latest) > MILLION_FORMAT_THRESHOLD
                        else f"{latest:,.2f}"
                    )
                    # Only show significant values (> $1M or shares > 1000)
                    if abs(latest) > SIGNIFICANT_VALUE_THRESHOLD or (
                        issue == 'shares' and latest > SHARES_SIGNIFICANT_THRESHOLD
                    ):
                        print(f"    {tag}: {val_str} (FY{years[0]})")


def main():
    # Investigate just the most problematic companies first
    priority = list(DEFAULT_PRIORITY)

    if len(sys.argv) > 1:
        priority = sys.argv[1:]

    for ticker in priority:
        if ticker in INVESTIGATE:
            investigate_company(ticker, INVESTIGATE[ticker])
        else:
            investigate_company(ticker, list(SEARCH_TAGS.keys()))
        import time
        time.sleep(SLEEP_SECONDS)


if __name__ == '__main__':
    main()
