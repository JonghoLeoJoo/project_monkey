"""
SEC EDGAR Data Fetcher
Fetches financial data from SEC EDGAR for publicly traded companies.
All financial values are returned in raw dollars (not thousands/millions).
"""

import requests
import time
from datetime import datetime, timedelta
from typing import Optional, Dict, Any, List, Tuple

try:
    import yfinance as yf
    _HAS_YFINANCE = True
except ImportError:
    _HAS_YFINANCE = False

HEADERS = {
    'User-Agent': 'FinancialModelApp research@example.com',
    'Accept': 'application/json',
}

TICKERS_URL = "https://www.sec.gov/files/company_tickers.json"
FACTS_BASE_URL = "https://data.sec.gov/api/xbrl/companyfacts"


def search_company(company_name: str, auto_select: bool = False) -> Optional[Dict]:
    """Search for a company by name or ticker. Returns {cik, name, ticker} or None.

    If auto_select is True, automatically picks the best match when multiple
    results are found (useful for non-interactive / batch mode).
    """
    print(f"  Searching SEC EDGAR for: '{company_name}'...")
    try:
        response = requests.get(TICKERS_URL, headers=HEADERS, timeout=30)
        response.raise_for_status()
        tickers = response.json()
    except requests.RequestException as e:
        print(f"  Error fetching company list: {e}")
        return None

    query = company_name.strip().lower()
    matches = []

    for _, company in tickers.items():
        title = company.get('title', '').lower()
        ticker = company.get('ticker', '').lower()
        if query == ticker or query in title or title.startswith(query):
            matches.append({
                'cik': str(company['cik_str']).zfill(10),
                'name': company['title'],
                'ticker': company['ticker'],
            })

    if not matches:
        print(f"  No company found for '{company_name}'.")
        return None

    if len(matches) == 1:
        print(f"  Found: {matches[0]['name']} ({matches[0]['ticker']}), CIK: {matches[0]['cik']}")
        return matches[0]

    # Sort by how closely name matches the query
    matches.sort(key=lambda m: (0 if m['ticker'].lower() == query else 1,
                                 0 if m['name'].lower() == query else 1,
                                 len(m['name'])))

    if auto_select:
        print(f"  Auto-selected: {matches[0]['name']} ({matches[0]['ticker']})")
        return matches[0]

    print(f"\n  Multiple matches found:")
    for i, m in enumerate(matches[:10], 1):
        print(f"    {i}. {m['name']} ({m['ticker']})")

    while True:
        try:
            choice = int(input(f"\n  Select company (1-{min(len(matches), 10)}): "))
            if 1 <= choice <= min(len(matches), 10):
                return matches[choice - 1]
        except (ValueError, KeyboardInterrupt):
            pass
        print("  Invalid choice, please try again.")


def get_company_facts(cik: str) -> Optional[Dict]:
    """Fetch all XBRL facts for a company from SEC EDGAR."""
    url = f"{FACTS_BASE_URL}/CIK{cik}.json"
    print(f"  Fetching XBRL data from SEC EDGAR (this may take a moment)...")
    try:
        time.sleep(0.1)  # respect SEC rate limits
        response = requests.get(url, headers=HEADERS, timeout=90)
        response.raise_for_status()
        return response.json()
    except requests.RequestException as e:
        print(f"  Error fetching company facts: {e}")
        return None


def _get_annual_values(facts: Dict, concept: str, taxonomy: str = 'us-gaap') -> List[Dict]:
    """Extract annual 10-K values for a XBRL concept. Returns list sorted newest-first."""
    try:
        concept_data = facts['facts'][taxonomy][concept]
        units = concept_data.get('units', {})
        for unit_key in ('USD', 'shares', 'pure'):
            if unit_key not in units:
                continue
            annual = [
                v for v in units[unit_key]
                if v.get('form') == '10-K' and v.get('fp') == 'FY'
                and 'end' in v and 'val' in v and v.get('fy')
            ]
            if not annual:
                continue
            # deduplicate: keep the latest filing for each fiscal year
            best = {}
            for v in annual:
                fy = v['fy']
                if fy not in best or v['end'] > best[fy]['end']:
                    best[fy] = v
            return sorted(best.values(), key=lambda x: x['end'], reverse=True)
        return []
    except (KeyError, TypeError):
        return []


def _get_metric(facts: Dict, concepts: List[str], years: List[int],
                taxonomy: str = 'us-gaap') -> Dict[int, Optional[float]]:
    """Try multiple XBRL concept names; return year->value mapping for the requested years."""
    result: Dict[int, Optional[float]] = {y: None for y in years}
    for concept in concepts:
        values = _get_annual_values(facts, concept, taxonomy)
        if not values:
            continue
        year_map = {v['fy']: v['val'] for v in values if v['fy'] in years}
        for y in years:
            if result[y] is None and y in year_map:
                result[y] = float(year_map[y])
        if all(result[y] is not None for y in years):
            break
    return result


def get_fiscal_years(facts: Dict, n_years: int = 3) -> List[int]:
    """Return the most recent n fiscal years available in the company facts."""
    probe_concepts = [
        'Assets', 'Revenues', 'NetIncomeLoss',
        'RevenueFromContractWithCustomerExcludingAssessedTax',
    ]
    years_found: set = set()
    for concept in probe_concepts:
        for v in _get_annual_values(facts, concept):
            if v.get('fy'):
                years_found.add(v['fy'])
    if not years_found:
        return []
    return sorted(years_found, reverse=True)[:n_years]


def extract_financial_data(facts: Dict, years: List[int]) -> Dict:
    """Extract all 3-statement financial data for the given fiscal years.

    Each metric tries an ordered list of XBRL concept names (most preferred first).
    The first concept that returns data for ALL requested years is used; if none
    covers every year, the function fills as many years as possible by combining
    results across concepts, so partial data is still captured.
    """
    print(f"  Extracting financial statements for FY{years[-1]}-FY{years[0]}...")

    def g(concepts, tax='us-gaap'):
        return _get_metric(facts, concepts, years, tax)

    # =========================================================================
    # INCOME STATEMENT
    # =========================================================================

    # Revenue – covers product, service, mixed, and legacy tags
    revenue = g([
        'Revenues',
        'RevenueFromContractWithCustomerExcludingAssessedTax',   # ASC 606 (most common modern)
        'RevenueFromContractWithCustomerIncludingAssessedTax',
        'SalesRevenueNet',                                        # legacy
        'SalesRevenueGoodsNet',
        'SalesRevenueServicesNet',
        'RevenuesNetOfInterestExpense',                           # banks / financial
        'InterestAndFeeIncomeLoansAndLeases',                     # banks
        'HealthCareOrganizationRevenue',                          # healthcare
        'RealEstateRevenueNet',                                   # real estate
        'RevenueFromContractWithCustomer',
    ])

    # Cost of goods / services sold
    cogs = g([
        'CostOfRevenue',
        'CostOfGoodsAndServicesSold',
        'CostOfGoodsSold',
        'CostOfServices',
        'CostOfGoodsAndServicesExcludingDepreciationDepletionAndAmortization',
        'CostOfRevenueExcludingDepreciationAndAmortization',
        'CostOfGoodsSoldExcludingDepreciationDepletionAndAmortization',
    ])

    gross_profit = g(['GrossProfit'])
    # Derive if not directly tagged
    for y in years:
        if gross_profit[y] is None and revenue[y] and cogs[y]:
            gross_profit[y] = revenue[y] - cogs[y]

    # Research & development
    rd_expense = g([
        'ResearchAndDevelopmentExpense',
        'ResearchAndDevelopmentExpenseExcludingAcquiredInProcessCost',
        'ResearchAndDevelopmentExpenseSoftwareExcludingAcquiredInProcessCost',
    ])

    # Selling, general & administrative
    # Some filers (e.g. Alphabet) report G&A and Sales & Marketing as separate
    # line items instead of a combined SGA figure.  Try the combined concept
    # first; if absent, sum the two components.
    sga = g(['SellingGeneralAndAdministrativeExpense'])
    if all(v is None for v in sga.values()):
        ga  = g(['GeneralAndAdministrativeExpense'])
        sm  = g(['SellingAndMarketingExpense', 'SellingExpense',
                 'MarketingAndAdvertisingExpense'])
        sga = {y: ((ga.get(y) or 0) + (sm.get(y) or 0))
                   if ga.get(y) is not None or sm.get(y) is not None
                   else None
               for y in years}

    # Operating income / EBIT
    operating_income = g([
        'OperatingIncomeLoss',
        'IncomeLossFromContinuingOperationsBeforeIncomeTaxesExtraordinaryItemsNoncontrollingInterest',
    ])

    # Depreciation & amortization (income-statement context)
    da = g([
        'DepreciationDepletionAndAmortization',
        'DepreciationAndAmortization',
        'Depreciation',
        'AmortizationOfIntangibleAssets',
        'DepreciationAmortizationAndAccretionNet',
        'OtherDepreciationAndAmortization',
        'DepreciationAndAmortizationExcludingDisposals',
    ])

    # EBITDA derived from EBIT + D&A
    ebitda: Dict[int, Optional[float]] = {}
    for y in years:
        if operating_income[y] is not None and da[y] is not None:
            ebitda[y] = operating_income[y] + da[y]
        else:
            ebitda[y] = None

    # Interest expense (gross)
    interest_expense = g([
        'InterestExpense',
        'InterestAndDebtExpense',
        'InterestExpenseDebt',
        'InterestExpenseRelatedParty',
        'FinanceLeaseInterestExpense',
        'InterestExpenseLongTermDebt',
        'InterestExpenseOther',
    ])

    # Interest / investment income
    interest_income = g([
        'InvestmentIncomeInterest',
        'InterestAndDividendIncomeOperating',
        'InterestIncomeOperating',
        'InterestIncomeOther',
        'InvestmentIncomeInterestAndDividend',
    ])

    # Other income / (expense) — non-operating items not captured by
    # interest income or interest expense (e.g. investment gains/losses,
    # FX gains/losses, equity method income).  Derived as:
    #   PreTax - EBIT - InterestIncome + InterestExpense
    # so that EBIT + IntInc - IntExp + OtherIncome = PreTax exactly.
    # Computed after pretax_income is extracted (see below).

    # Net non-operating income — fallback for filers that report a single
    # net figure instead of separate interest income / expense tags
    # (e.g. Apple FY2024+).
    nonop_net = g([
        'NonoperatingIncomeExpense',
        'OtherNonoperatingIncomeExpense',
    ])

    # Pre-tax income
    pretax_income = g([
        'IncomeLossFromContinuingOperationsBeforeIncomeTaxesExtraordinaryItemsNoncontrollingInterest',
        'IncomeLossFromContinuingOperationsBeforeIncomeTaxesMinorityInterestAndIncomeLossFromEquityMethodInvestments',
        'IncomeLossFromContinuingOperationsBeforeIncomeTaxesDomestic',
    ])

    # Other income / (expense): plug so that EBIT + IntInc - IntExp + Other = PreTax
    # For years where interest expense & income are both missing but
    # NonoperatingIncomeExpense exists, use the net figure directly as
    # "other income" and set interest items to 0 so the IS still balances.
    other_income = {}
    for yr in years:
        pt = pretax_income.get(yr)
        ebit = operating_income.get(yr)
        if pt is None or ebit is None:
            other_income[yr] = None
            continue

        ie = interest_expense.get(yr)
        ii = interest_income.get(yr)

        if ie is None and ii is None and nonop_net.get(yr) is not None:
            # Filer consolidated all non-operating items into one net line.
            # Use NonoperatingIncomeExpense as "other income" and zero out
            # interest so the IS balances: EBIT + 0 - 0 + NonopNet = PreTax.
            interest_expense[yr] = 0.0
            interest_income[yr] = 0.0
            other_income[yr] = nonop_net[yr]
        else:
            other_income[yr] = pt - ebit - (ii or 0) + (ie or 0)

    # Income tax expense
    tax_expense = g([
        'IncomeTaxExpenseBenefit',
        'CurrentIncomeTaxExpenseBenefit',
        'IncomeTaxesPaidNet',
    ])

    # Net income
    net_income = g([
        'NetIncomeLoss',
        'NetIncome',
        'ProfitLoss',
        'NetIncomeLossAvailableToCommonStockholdersBasic',
        'NetIncomeLossAttributableToParent',
        'IncomeLossFromContinuingOperations',
        'IncomeLossFromContinuingOperationsIncludingPortionAttributableToNoncontrollingInterest',
    ])

    # Earnings per share
    eps_basic = g([
        'EarningsPerShareBasic',
        'IncomeLossFromContinuingOperationsPerBasicShare',
        'EarningsPerShareBasicAndDiluted',
    ])
    eps_diluted = g([
        'EarningsPerShareDiluted',
        'IncomeLossFromContinuingOperationsPerDilutedShare',
        'EarningsPerShareBasicAndDiluted',
    ])

    # Shares outstanding
    shares_basic = g([
        'WeightedAverageNumberOfSharesOutstandingBasic',
        'CommonStockSharesOutstanding',
        'CommonStockSharesIssued',
    ])
    shares_diluted = g([
        'WeightedAverageNumberOfDilutedSharesOutstanding',
        'WeightedAverageNumberOfSharesOutstandingDiluted',
        'CommonStockSharesOutstanding',
    ])

    # =========================================================================
    # BALANCE SHEET – ASSETS
    # =========================================================================

    # Cash & cash equivalents (incl. restricted cash when present).
    # Prefer the CF-statement reconciling concept so that the CF reconciliation
    # section ties to this balance. For companies with no restricted cash both
    # concepts return the same value.
    cash = g([
        'CashCashEquivalentsRestrictedCashAndRestrictedCashEquivalents',
        'CashCashEquivalentsRestrictedCashAndRestrictedCashEquivalentsIncludingDisposalGroupAndDiscontinuedOperations',
        'CashAndCashEquivalentsAtCarryingValue',
        'CashAndCashEquivalents',
        'Cash',
        'CashAndDueFromBanks',                          # banks
        'CashEquivalentsAtCarryingValue',
        'RestrictedCashAndCashEquivalentsAtCarryingValue',
    ])

    # Short-term / marketable securities
    st_investments = g([
        'ShortTermInvestments',
        'MarketableSecuritiesCurrent',
        'AvailableForSaleSecuritiesCurrent',
        'DebtSecuritiesAvailableForSaleCurrent',
        'TradingSecuritiesCurrent',
        'HeldToMaturitySecuritiesCurrent',
        'EquitySecuritiesFvNiCurrent',
        'InvestmentsCurrent',
    ])

    # Accounts receivable
    accounts_rec = g([
        'AccountsReceivableNetCurrent',
        'ReceivablesNetCurrent',
        'TradeAndOtherReceivablesNetCurrent',
        'AccountsAndNotesReceivableNet',
        'AccountsReceivableNet',
        'ContractWithCustomerAssetNetCurrent',          # ASC 606 contract asset
        'BilledContractReceivables',
        'AccountsReceivableGrossCurrent',
    ])

    # Inventory – most industry variants covered
    inventory = g([
        'InventoryNet',
        'Inventories',
        'InventoryGross',
        'InventoryFinishedGoods',
        'InventoryFinishedGoodsNetOfReserves',
        'InventoryWorkInProcess',
        'InventoryRawMaterials',
        'InventoryRawMaterialsNetOfReserves',
        'InventoryAndOther',                            # catch-all some filers use
        'SuppliesAndInventories',
        'RetailRelatedInventoryMerchandise',            # retailers
        'FuelInventory',                                # energy / airlines
        'InventoryNetOfAllowancesCustomerAdvancesAndProgress',
        'InventoryRealEstate',                          # homebuilders / real estate
        'AircraftMaintenanceMaterialsAndRepairs',       # airlines
    ])

    # Other current assets / prepaid expenses
    other_current_a = g([
        'OtherAssetsCurrent',
        'PrepaidExpenseAndOtherAssetsCurrent',
        'PrepaidExpenseCurrent',
        'OtherCurrentAssets',
        'DeferredCostsCurrent',
        'PrepaidExpenseOtherCurrent',
        'AssetsCurrentOther',
    ])

    total_current_a = g(['AssetsCurrent'])

    # PP&E, net
    ppe_net = g([
        'PropertyPlantAndEquipmentNet',
        'PropertyPlantAndEquipmentNetExcludingCapitalizedInterest',
        'PropertyAndEquipmentNet',
        'PropertyPlantAndEquipmentAndRightOfUseAssetAfterAccumulatedDepreciationAndAmortization',
        # Amazon (and others post-ASC 842): PP&E + finance lease ROU assets combined
        'PropertyPlantAndEquipmentAndFinanceLeaseRightOfUseAssetAfterAccumulatedDepreciationAndAmortization',
        'PropertyPlantEquipmentAndFinanceLeaseRightOfUseAssetNet',
    ])

    # Goodwill
    goodwill = g([
        'Goodwill',
        'GoodwillGross',
    ])

    # Intangible assets (ex-goodwill)
    intangibles = g([
        'IntangibleAssetsNetExcludingGoodwill',
        'FiniteLivedIntangibleAssetsNet',
        'IndefiniteLivedIntangibleAssetsExcludingGoodwill',
        'IntangibleAssetsNet',
        'OtherIntangibleAssetsNet',
        'CapitalizedComputerSoftwareNet',
    ])

    # Marketable securities (non-current) / long-term investments
    lt_investments = g([
        'MarketableSecuritiesNoncurrent',
        'AvailableForSaleSecuritiesNoncurrent',
        'OtherLongTermInvestments',
        'LongTermInvestments',
        'InvestmentsAndOtherNoncurrentAssets',
        'EquitySecuritiesWithoutReadilyDeterminableFairValueAmount',
        'EquityMethodInvestments',
    ])

    # Other non-current assets
    other_noncurrent_a = g([
        'OtherAssetsNoncurrent',
        'OtherNoncurrentAssets',
        'OtherLongTermAssets',
        'DeferredIncomeTaxAssetsNet',
        'OperatingLeaseRightOfUseAsset',                # ASC 842 ROU (if not in PPE)
        'NoncurrentAssets',
    ])

    total_assets = g(['Assets'])

    # =========================================================================
    # BALANCE SHEET – LIABILITIES
    # =========================================================================

    # Accounts payable
    accounts_pay = g([
        'AccountsPayableCurrent',
        'AccountsPayable',
        'AccountsPayableAndAccruedLiabilitiesCurrent',
        'TradeAccountsPayableCurrent',
        'AccountsPayableRelatedPartiesCurrent',
    ])

    # Accrued liabilities (employee-related, accrued expenses, etc.)
    accrued_liab = g([
        'AccruedLiabilitiesCurrent',
        'EmployeeRelatedLiabilitiesCurrent',
        'AccruedAndOtherCurrentLiabilities',
        'AccruedExpensesAndOtherCurrentLiabilities',
        'OtherAccruedLiabilitiesCurrent',
        'AccruedExpensesCurrent',
    ])

    # Other current liabilities (catch-all for CL items not in the slots above)
    other_current_l = g([
        'OtherLiabilitiesCurrent',
        'OtherCurrentLiabilities',
        'AccruedIncomeTaxesCurrent',
        'TaxesPayableCurrent',
        'OtherCurrentLiabilitiesAndAccruedExpenses',
    ])

    # Short-term / current portion of long-term debt
    st_debt = g([
        'ShortTermBorrowings',
        'LongTermDebtCurrent',
        'CurrentMaturitiesOfLongTermDebt',
        'DebtCurrent',
        'NotesPayableCurrent',
        'ShortTermDebtAndCurrentMaturitiesOfLongTermDebt',
        'LongTermDebtAndCapitalLeaseObligationsCurrent',
        'CommercialPaper',
        'LinesOfCreditCurrent',
    ])

    # Deferred revenue (current)
    deferred_rev_cur = g([
        'DeferredRevenueCurrent',
        'ContractWithCustomerLiabilityCurrent',
        'DeferredRevenueAndCredits',
        'CustomerAdvancesCurrent',
        'DeferredIncomeCurrent',
    ])

    total_current_l = g(['LiabilitiesCurrent'])

    # Long-term debt
    lt_debt = g([
        'LongTermDebtNoncurrent',
        'LongTermDebt',
        'LongTermDebtExcludingCurrentMaturities',
        'LongTermDebtAndCapitalLeaseObligations',
        'LongTermNotesPayable',
        'SeniorLongTermNotes',
        'ConvertibleDebtNoncurrent',
        'UnsecuredLongTermDebt',
        'LongTermBorrowings',
    ])

    # Deferred tax liabilities / income taxes payable (non-current)
    deferred_tax_l = g([
        'DeferredIncomeTaxLiabilitiesNet',
        'DeferredTaxLiabilitiesNoncurrent',
        'DeferredTaxLiabilities',
        'DeferredIncomeTaxesAndOtherTaxLiabilitiesNoncurrent',
        'DeferredTaxAndOtherLiabilitiesNoncurrent',
        'AccruedIncomeTaxesNoncurrent',
    ])

    # Other non-current liabilities
    other_noncurrent_l = g([
        'OtherLiabilitiesNoncurrent',
        'OtherNoncurrentLiabilities',
        'OtherLongTermLiabilities',
        'LiabilitiesOtherThanLongtermDebtNoncurrent',
        'OperatingLeaseLiabilityNoncurrent',            # ASC 842 operating lease
    ])

    total_liabilities = g([
        'Liabilities',
        # Note: LiabilitiesNoncurrent intentionally excluded — it omits current
        # liabilities and would break the A = L + E identity.  When 'Liabilities'
        # is not tagged, the arithmetic fallback (Assets - Equity) below derives
        # the correct total.
    ])

    # =========================================================================
    # BALANCE SHEET – EQUITY
    # =========================================================================

    common_stock = g([
        'CommonStockValue',
        'CommonStockValueOutstanding',
        'CommonStocksIncludingAdditionalPaidInCapital',  # some filers combine stock + APIC
    ])

    apic = g([
        'AdditionalPaidInCapital',
        'AdditionalPaidInCapitalCommonStock',
    ])

    retained_earnings = g([
        'RetainedEarningsAccumulatedDeficit',
        'RetainedEarnings',
        'RetainedEarningsAppropriated',
    ])

    treasury_stock = g([
        'TreasuryStockValue',
        'TreasuryStockCommonValue',
        'TreasuryStockAtCost',
    ])

    total_equity = g([
        # Prefer the concept that includes noncontrolling interests (NCI / minority interest)
        # so the accounting identity Assets = Liabilities + Total Equity holds.
        'StockholdersEquityIncludingPortionAttributableToNoncontrollingInterest',
        'StockholdersEquity',
        'PartnersCapital',
        'MembersEquity',
        'LimitedLiabilityCompanyLlcMembersEquityIncludingPortionAttributableToNoncontrollingInterest',
    ])

    # Total Liabilities (plug): always derived as Assets - Equity so that the
    # balance-sheet identity A = L + E holds by construction.  This absorbs
    # redeemable NCI / mezzanine equity automatically, regardless of whether
    # the filer tags "Liabilities" as a standalone concept.
    for yr in years:
        ta = total_assets.get(yr)
        te = total_equity.get(yr)
        if ta is not None and te is not None:
            total_liabilities[yr] = ta - te

    # =========================================================================
    # CASH FLOW STATEMENT
    # =========================================================================

    # Operating cash flow
    operating_cf = g([
        'NetCashProvidedByUsedInOperatingActivities',
        'NetCashProvidedByUsedInOperatingActivitiesContinuingOperations',
        'CashGeneratedFromOperations',
    ])

    # Capital expenditures (always reported as a positive outflow in XBRL)
    capex = g([
        'PaymentsToAcquirePropertyPlantAndEquipment',
        'AcquisitionsOfPropertyPlantAndEquipment',
        'PurchasesOfPropertyAndEquipment',
        'PaymentsForCapitalImprovements',
        'PaymentsToAcquirePropertyPlantEquipmentAndOther',
        'PaymentsToAcquireProductiveAssets',             # Amazon et al.
        'InvestmentsInPropertyPlantAndEquipment',
        'PaymentsToAcquireAndDevelopRealEstate',         # real estate / homebuilders
        'CapitalExpenditureDiscontinuedOperations',
    ])

    # Business acquisitions
    acquisitions = g([
        'PaymentsToAcquireBusinessesNetOfCashAcquired',
        'PaymentsToAcquireBusinessesGross',
        'PaymentsToAcquireBusinessesAndInterestInAffiliates',
        'BusinessAcquisitionCostOfAcquiredEntityTransactionCosts',
    ])

    # Investing cash flow
    investing_cf = g([
        'NetCashProvidedByUsedInInvestingActivities',
        'NetCashProvidedByUsedInInvestingActivitiesContinuingOperations',
    ])

    # Dividends paid
    dividends = g([
        'PaymentsOfDividends',
        'PaymentsOfDividendsCommonStock',
        'PaymentsOfOrdinaryDividends',
        'DividendsPaid',
        'PaymentsOfDividendsAndDividendEquivalentsOnRestrictedStockUnitsAndOtherEquityAwards',
    ])

    # Share repurchases
    repurchases = g([
        'PaymentsForRepurchaseOfCommonStock',
        'TreasuryStockValueAcquiredCostMethod',
        'PaymentsForRepurchaseOfEquity',
        'PaymentsForRepurchaseOfCommonStockAndStockAwards',
        'RepurchaseOfCommonStock',
    ])

    # Debt issuance proceeds
    debt_issuance = g([
        'ProceedsFromIssuanceOfLongTermDebt',
        'ProceedsFromDebtNetOfIssuanceCosts',
        'ProceedsFromIssuanceOfDebt',
        'ProceedsFromBorrowings',
        'ProceedsFromLongTermLinesOfCredit',
        'ProceedsFromIssuanceOfSeniorLongTermDebt',
        'ProceedsFromIssuanceOfUnsecuredDebt',
    ])

    # Debt repayment
    debt_repay = g([
        'RepaymentsOfLongTermDebt',
        'RepaymentsOfDebt',
        'RepaymentsOfBorrowings',
        'RepaymentOfLongTermDebt',
        'RepaymentsOfSeniorDebt',
        'RepaymentsOfLinesOfCredit',
        'RepaymentsOfUnsecuredDebt',
        'RepaymentsOfNotesPayable',
    ])

    # Financing cash flow
    financing_cf = g([
        'NetCashProvidedByUsedInFinancingActivities',
        'NetCashProvidedByUsedInFinancingActivitiesContinuingOperations',
    ])

    # Effect of exchange rate changes on cash
    # This 4th CF category reconciles Op+Inv+Fin totals to the actual BS cash change.
    fx_effect = g([
        'EffectOfExchangeRateOnCashCashEquivalentsRestrictedCashAndRestrictedCashEquivalentsIncludingDisposalGroupAndDiscontinuedOperations',
        'EffectOfExchangeRateOnCashCashEquivalentsRestrictedCashAndRestrictedCashEquivalents',
        'EffectOfExchangeRateOnCashAndCashEquivalents',
        'EffectOfExchangeRateOnCashAndCashEquivalentsContinuingOperations',
        'ExchangeRateEffectOnCashAndCashEquivalents',
        'EffectOfExchangeRateOnCash',
    ])

    # Stock-based compensation (non-cash add-back in operating CF)
    sbc = g([
        'ShareBasedCompensation',
        'AllocatedShareBasedCompensationExpense',
        'ShareBasedCompensationExpense',
        'EmployeeBenefitsAndShareBasedCompensation',
        'StockBasedCompensation',
    ])

    fcf: Dict[int, Optional[float]] = {}
    for y in years:
        op = operating_cf[y]
        cx = capex[y]
        if op is not None and cx is not None:
            fcf[y] = op - abs(cx)
        else:
            fcf[y] = None

    result = {
        'years': years,
        'income_statement': {
            'revenue':          revenue,
            'cogs':             cogs,
            'gross_profit':     gross_profit,
            'rd_expense':       rd_expense,
            'sga_expense':      sga,
            'operating_income': operating_income,
            'ebitda':           ebitda,
            'da':               da,
            'interest_expense': interest_expense,
            'interest_income':  interest_income,
            'other_income':     other_income,
            'pretax_income':    pretax_income,
            'tax_expense':      tax_expense,
            'net_income':       net_income,
            'eps_basic':        eps_basic,
            'eps_diluted':      eps_diluted,
            'shares_basic':     shares_basic,
            'shares_diluted':   shares_diluted,
        },
        'balance_sheet': {
            'cash':               cash,
            'st_investments':     st_investments,
            'accounts_rec':       accounts_rec,
            'inventory':          inventory,
            'other_current_a':    other_current_a,
            'total_current_a':    total_current_a,
            'ppe_net':            ppe_net,
            'goodwill':           goodwill,
            'intangibles':        intangibles,
            'lt_investments':     lt_investments,
            'other_noncurrent_a': other_noncurrent_a,
            'total_assets':       total_assets,
            'accounts_pay':       accounts_pay,
            'accrued_liab':       accrued_liab,
            'other_current_l':    other_current_l,
            'st_debt':            st_debt,
            'deferred_rev_cur':   deferred_rev_cur,
            'total_current_l':    total_current_l,
            'lt_debt':            lt_debt,
            'deferred_tax_l':     deferred_tax_l,
            'other_noncurrent_l': other_noncurrent_l,
            'total_liabilities':  total_liabilities,
            'common_stock':       common_stock,
            'apic':               apic,
            'retained_earnings':  retained_earnings,
            'treasury_stock':     treasury_stock,
            'total_equity':       total_equity,
        },
        'cash_flow': {
            'net_income':   net_income,
            'da':           da,
            'sbc':          sbc,
            'operating_cf': operating_cf,
            'capex':        capex,
            'acquisitions': acquisitions,
            'investing_cf': investing_cf,
            'dividends':    dividends,
            'repurchases':  repurchases,
            'debt_issuance': debt_issuance,
            'debt_repay':   debt_repay,
            'financing_cf': financing_cf,
            'fx_effect':    fx_effect,
            'fcf':          fcf,
        },
    }

    # Extended revenue history for 5-year average growth calculation.
    # Fetch up to 6 years of revenue (gives 5 YoY growth periods).
    all_fy = sorted(get_fiscal_years(facts, n_years=max(6, len(years))),
                    reverse=True)
    extended_rev = _get_metric(facts, [
        'Revenues',
        'RevenueFromContractWithCustomerExcludingAssessedTax',
        'RevenueFromContractWithCustomerIncludingAssessedTax',
        'SalesRevenueNet',
        'SalesRevenueGoodsNet',
        'SalesRevenueServicesNet',
        'RevenuesNetOfInterestExpense',
        'InterestAndFeeIncomeLoansAndLeases',
        'HealthCareOrganizationRevenue',
        'RealEstateRevenueNet',
        'RevenueFromContractWithCustomer',
    ], all_fy)
    result['revenue_extended'] = extended_rev

    return result


def get_fiscal_year_end_dates(facts: Dict, years: List[int]) -> Dict[int, Optional[str]]:
    """Extract the exact fiscal year-end date (YYYY-MM-DD) per FY from XBRL facts.

    Uses the 'end' field from balance-sheet concepts (point-in-time), which
    represents the period-end date in the 10-K filing.
    """
    result: Dict[int, Optional[str]] = {y: None for y in years}
    for concept in ('Assets', 'Liabilities',
                    'StockholdersEquityIncludingPortionAttributableToNoncontrollingInterest',
                    'StockholdersEquity'):
        for v in _get_annual_values(facts, concept):
            fy = v.get('fy')
            if fy in result and result[fy] is None and 'end' in v:
                result[fy] = v['end']
        if all(result[y] is not None for y in years):
            break
    return result


def get_historical_closing_prices(
    ticker: str,
    fy_end_dates: Dict[int, Optional[str]],
) -> Dict[int, Optional[float]]:
    """Fetch historical closing stock prices at each fiscal year-end date via yfinance.

    Returns {fiscal_year: closing_price} with None for any year that could not be
    resolved.  If yfinance is not installed the function returns all-None gracefully.
    """
    result: Dict[int, Optional[float]] = {y: None for y in fy_end_dates}

    if not _HAS_YFINANCE:
        print("  Warning: yfinance not installed; skipping stock price fetch.")
        print("           Install with: pip install yfinance")
        return result

    try:
        stock = yf.Ticker(ticker)
    except Exception as e:
        print(f"  Warning: yfinance could not initialise ticker '{ticker}': {e}")
        return result

    for fy, date_str in fy_end_dates.items():
        if date_str is None:
            continue
        try:
            target = datetime.strptime(date_str, '%Y-%m-%d')
            # 10-day lookback window handles weekends / holidays;
            # end is exclusive in yfinance, so add 1 day.
            start = (target - timedelta(days=10)).strftime('%Y-%m-%d')
            end   = (target + timedelta(days=1)).strftime('%Y-%m-%d')

            hist = stock.history(start=start, end=end, auto_adjust=True)

            if hist.empty:
                print(f"  Warning: no price data for {ticker} near {date_str} (FY{fy})")
                continue

            closing_price = float(hist['Close'].iloc[-1])
            result[fy] = closing_price
            print(f"    FY{fy} ({date_str}): ${closing_price:.2f}")
        except Exception as e:
            print(f"  Warning: could not fetch price for FY{fy} ({date_str}): {e}")

    return result


def get_current_price(ticker: str, target_date: Optional[str] = None) -> Dict[str, Any]:
    """Fetch a closing stock price via yfinance.

    If *target_date* is given (``'YYYY-MM-DD'``), the function looks up the
    closing price on or just before that date (10-day lookback).  Otherwise
    it returns the most recent available close.

    Returns ``{'price': float|None, 'date': str|None}``.
    """
    if not _HAS_YFINANCE:
        return {'price': None, 'date': None}
    try:
        stock = yf.Ticker(ticker)
        if target_date:
            target = datetime.strptime(target_date, '%Y-%m-%d')
            start = (target - timedelta(days=10)).strftime('%Y-%m-%d')
            end   = (target + timedelta(days=1)).strftime('%Y-%m-%d')
            hist = stock.history(start=start, end=end, auto_adjust=True)
        else:
            hist = stock.history(period='5d')
        if not hist.empty:
            price = float(hist['Close'].iloc[-1])
            date_str = hist.index[-1].strftime('%Y-%m-%d')
            print(f"  Price for {ticker}: ${price:.2f} ({date_str})")
            return {'price': price, 'date': date_str}
    except Exception as e:
        print(f"  Warning: could not fetch price for {ticker}: {e}")
    return {'price': None, 'date': None}


def get_treasury_yield() -> Optional[float]:
    """Fetch the 10-Year US Treasury yield from CNBC, with yfinance fallback.

    Primary source: CNBC REST API (https://www.cnbc.com/quotes/US10Y).
    Fallback: yfinance ^TNX ticker.

    Returns the yield as a decimal (e.g. 0.045 for 4.5%) or *None*.
    """
    # Primary: CNBC REST API
    try:
        cnbc_url = ("https://quote.cnbc.com/quote-html-webservice/restQuote/"
                    "symbolType/symbol?symbols=US10Y&requestMethod=itv"
                    "&noCache=1&output=json")
        resp = requests.get(cnbc_url, timeout=10)
        resp.raise_for_status()
        data = resp.json()
        # Navigate JSON: {"FormattedQuoteResult":{"FormattedQuote":[{"last":"4.25",...}]}}
        quotes = data.get('FormattedQuoteResult', {}).get('FormattedQuote', [])
        if quotes:
            last_str = quotes[0].get('last', '')
            yield_pct = float(last_str)
            print(f"  10Y Treasury yield (CNBC): {yield_pct:.2f}%")
            return yield_pct / 100.0
    except Exception as e:
        print(f"  Warning: CNBC API unavailable ({e}), trying yfinance...")

    # Fallback: yfinance ^TNX
    if not _HAS_YFINANCE:
        return None
    try:
        tnx = yf.Ticker('^TNX')
        hist = tnx.history(period='5d')
        if not hist.empty:
            yield_pct = float(hist['Close'].iloc[-1])
            print(f"  10Y Treasury yield (yfinance): {yield_pct:.2f}%")
            return yield_pct / 100.0
    except Exception as e:
        print(f"  Warning: could not fetch treasury yield: {e}")
    return None


def get_kroll_erp() -> Optional[float]:
    """Fetch Kroll's recommended US Equity Risk Premium from kroll.com.

    Scrapes https://www.kroll.com/en/reports/cost-of-capital/
    recommended-us-equity-risk-premium-and-corresponding-risk-free-rates
    for the most recent ERP recommendation.

    Returns the ERP as a decimal (e.g. 0.05 for 5.0%) or *None*.
    """
    import re
    try:
        kroll_url = ("https://www.kroll.com/en/reports/cost-of-capital/"
                     "recommended-us-equity-risk-premium-and-corresponding-"
                     "risk-free-rates")
        resp = requests.get(kroll_url, headers={
            'User-Agent': ('Mozilla/5.0 (Windows NT 10.0; Win64; x64) '
                           'AppleWebKit/537.36 (KHTML, like Gecko) '
                           'Chrome/120.0.0.0 Safari/537.36'),
        }, timeout=15)
        resp.raise_for_status()
        html = resp.text
        # Look for the most recent ERP announcement in the HTML, e.g.:
        # "Equity Risk Premium to 5.0%, Effective June 5, 2024"
        matches = re.findall(
            r'Equity Risk Premium[^%]*?(\d+\.?\d*)%', html, re.IGNORECASE)
        if matches:
            erp_pct = float(matches[0])
            print(f"  Kroll Recommended ERP: {erp_pct:.1f}%")
            return erp_pct / 100.0
    except Exception as e:
        print(f"  Warning: could not fetch Kroll ERP ({e})")
    return None


def get_comparable_data(ticker: str) -> Optional[Dict]:
    """Fetch comparable company data from yfinance for WACC beta calculation.

    Returns a dict with keys: *name*, *ticker*, *beta*, *price*, *shares*
    (raw count), *total_debt* (raw $), *total_cash* (raw $), *net_debt*
    (raw $), *tax_rate* (decimal).  Returns ``None`` if critical fields
    (beta, price, shares) are missing.
    """
    if not _HAS_YFINANCE:
        return None
    try:
        stock = yf.Ticker(ticker)
        info = stock.info

        name   = info.get('shortName') or info.get('longName') or ticker
        beta   = info.get('beta')
        price  = info.get('currentPrice') or info.get('regularMarketPrice')
        shares = info.get('sharesOutstanding')  # raw number

        total_debt = info.get('totalDebt') or 0
        total_cash = info.get('totalCash') or 0

        # Tax rate from income statement
        tax_rate = None
        try:
            inc = stock.income_stmt
            if inc is not None and not inc.empty:
                tax_prov = inc.loc['Tax Provision'].iloc[0] if 'Tax Provision' in inc.index else None
                pretax   = inc.loc['Pretax Income'].iloc[0] if 'Pretax Income' in inc.index else None
                if tax_prov is not None and pretax is not None and pretax > 0:
                    tax_rate = float(tax_prov / pretax)
        except Exception:
            pass

        if beta is None or price is None or shares is None:
            print(f"  Warning: incomplete data for {ticker} (beta={beta}, price={price}, shares={shares})")
            return None

        return {
            'name':       name,
            'ticker':     ticker,
            'beta':       float(beta),
            'price':      float(price),
            'shares':     float(shares),
            'total_debt': float(total_debt),
            'total_cash': float(total_cash),
            'net_debt':   float(total_debt - total_cash),
            'tax_rate':   float(tax_rate) if tax_rate is not None else 0.21,
        }
    except Exception as e:
        print(f"  Warning: could not fetch comparable data for {ticker}: {e}")
        return None


def get_industry_peers(ticker: str, max_peers: int = 10) -> Tuple[List[Dict], str]:
    """Find comparable companies in the same industry using yfinance.

    Uses ``yf.Industry`` to find peers in the same Yahoo Finance industry
    classification.  Falls back to ``yf.Sector`` if the industry has fewer
    than *max_peers* candidates.

    Returns ``(comp_data_list, industry_name)`` where *comp_data_list*
    contains up to *max_peers* dicts (same schema as
    :func:`get_comparable_data`) and *industry_name* is the human-readable
    industry string (e.g. ``"Internet Retail"``).
    """
    if not _HAS_YFINANCE:
        return [], ''

    try:
        stock = yf.Ticker(ticker)
        info = stock.info
        industry_key = info.get('industryKey')
        sector_key = info.get('sectorKey')
        industry_name = info.get('industry', '')

        if not industry_key:
            return [], industry_name

        # Gather candidate tickers from industry, then sector as fallback
        candidates: List[str] = []
        use_sector = False

        try:
            ind = yf.Industry(industry_key)
            tc = ind.top_companies
            if tc is not None and not tc.empty:
                # If the target dominates its industry (>85% weight),
                # the remaining peers are too small to be meaningful comps.
                if 'market weight' in tc.columns and ticker in tc.index:
                    tgt_wt = tc.loc[ticker, 'market weight']
                    if tgt_wt is not None and tgt_wt > 0.90:
                        use_sector = True
                if not use_sector:
                    candidates = [sym for sym in tc.index
                                  if sym != ticker and isinstance(sym, str)]
        except Exception:
            pass

        # If industry too narrow or target dominates, use sector peers
        if (use_sector or len(candidates) < max_peers) and sector_key:
            try:
                sec = yf.Sector(sector_key)
                tc = sec.top_companies
                if tc is not None and not tc.empty:
                    seen = set(candidates) | {ticker}
                    extras = [sym for sym in tc.index
                              if sym not in seen and isinstance(sym, str)]
                    candidates.extend(extras)
            except Exception:
                pass

        # Fetch comparable data for top candidates until we have enough
        results: List[Dict] = []
        for ct in candidates:
            if len(results) >= max_peers:
                break
            cd = get_comparable_data(ct)
            if cd:
                results.append(cd)

        return results, industry_name

    except Exception as e:
        print(f"  Warning: could not find industry peers for {ticker}: {e}")
        return [], ''
