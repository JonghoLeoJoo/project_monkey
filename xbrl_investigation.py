import requests, json, time

HEADERS = {'User-Agent': 'research@test.com', 'Accept-Encoding': 'gzip, deflate'}

def get_xbrl_facts(ticker):
    """Get company's XBRL data from SEC EDGAR."""
    # Search for CIK
    url = f"https://efts.sec.gov/LATEST/search-index?q=%22{ticker}%22&dateRange=custom&startdt=2024-01-01&enddt=2025-12-31&forms=10-K"
    resp = requests.get(f"https://efts.sec.gov/LATEST/company-search?company={ticker}&CIK=&type=10-K&dateb=&owner=include&count=5&search_text=&action=getcompany",
                       headers=HEADERS)

    # Try direct ticker lookup
    tickers_url = "https://www.sec.gov/files/company_tickers.json"
    tickers_resp = requests.get(tickers_url, headers=HEADERS)
    tickers_data = tickers_resp.json()

    cik = None
    for entry in tickers_data.values():
        if entry.get('ticker', '').upper() == ticker.upper():
            cik = str(entry['cik_str']).zfill(10)
            break

    if not cik:
        print(f"  Could not find CIK for {ticker}")
        return None

    facts_url = f"https://data.sec.gov/api/xbrl/companyfacts/CIK{cik}.json"
    resp = requests.get(facts_url, headers=HEADERS)
    return resp.json()

def investigate_company(ticker, search_terms):
    """Search for specific concepts in a company's XBRL data."""
    print(f"\n{'='*60}")
    print(f"  {ticker} - XBRL Investigation")
    print(f"{'='*60}")

    facts = get_xbrl_facts(ticker)
    if not facts:
        return

    for taxonomy in ['us-gaap', 'dei']:
        if taxonomy not in facts.get('facts', {}):
            continue
        concepts = facts['facts'][taxonomy]
        for term in search_terms:
            matches = [(k, v) for k, v in concepts.items()
                       if term.lower() in k.lower()]
            if matches:
                print(f"\n  [{taxonomy}] Matches for '{term}':")
                for name, data in matches:
                    units = data.get('units', {})
                    for unit_type, entries in units.items():
                        # Get annual entries (10-K)
                        annual = [e for e in entries if e.get('form') == '10-K'
                                  and 'frame' in e and e['frame'].startswith('CY202')]
                        if annual:
                            print(f"    {name} ({unit_type}):")
                            for e in annual[-5:]:
                                val = e.get('val', 'N/A')
                                if isinstance(val, (int, float)):
                                    val = f"${val/1e6:,.0f}M" if abs(val) > 100000 else f"{val:,.2f}"
                                print(f"      {e.get('frame','?')}: {val} (filed {e.get('filed','?')})")
    time.sleep(0.2)

# 1. AAPL - What interest expense tag does Apple use?
investigate_company('AAPL', ['Interest', 'Debt'])

time.sleep(0.5)

# 2. ORCL - Shares investigation
investigate_company('ORCL', ['Share', 'CommonStock', 'EarningsPerShare'])

time.sleep(0.5)

# 3. AMGN - Why is WACC so low? Check debt/equity structure
investigate_company('AMGN', ['Debt', 'LongTermDebt', 'Interest', 'Beta'])

print("\nDone!")
