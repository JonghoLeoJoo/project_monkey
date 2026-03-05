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


# =========================================================================
# CONCEPT LISTS  –  shared between extract_financial_data & compute_ltm
# =========================================================================

IS_CONCEPTS: Dict[str, List[str]] = {
    'revenue': [
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
    ],
    'cogs': [
        'CostOfRevenue',
        'CostOfGoodsAndServicesSold',
        'CostOfGoodsSold',
        'CostOfServices',
        'CostOfGoodsAndServicesExcludingDepreciationDepletionAndAmortization',
        'CostOfGoodsAndServiceExcludingDepreciationDepletionAndAmortization',
        'CostOfRevenueExcludingDepreciationAndAmortization',
        'CostOfGoodsSoldExcludingDepreciationDepletionAndAmortization',
    ],
    'gross_profit': ['GrossProfit'],
    'rd_expense': [
        'ResearchAndDevelopmentExpense',
        'ResearchAndDevelopmentExpenseExcludingAcquiredInProcessCost',
        'ResearchAndDevelopmentExpenseSoftwareExcludingAcquiredInProcessCost',
    ],
    'sga_expense': ['SellingGeneralAndAdministrativeExpense'],
    'sga_ga': ['GeneralAndAdministrativeExpense'],
    'sga_sm': ['SellingAndMarketingExpense', 'SellingExpense',
               'MarketingAndAdvertisingExpense'],
    'operating_income': ['OperatingIncomeLoss'],
    'da': [
        'DepreciationDepletionAndAmortization',
        'DepreciationAndAmortization',
        'Depreciation',
        'AmortizationOfIntangibleAssets',
        'DepreciationAmortizationAndAccretionNet',
        'OtherDepreciationAndAmortization',
        'DepreciationAndAmortizationExcludingDisposals',
    ],
    'amortization': [
        'AmortizationOfIntangibleAssets',
        'Amortization',
    ],
    'transformation_costs': [
        'BusinessCombinationIntegrationRelatedCosts',
        'BusinessCombinationAcquisitionRelatedCosts',
        'RestructuringCharges',
        'RestructuringAndRelatedCostIncurredCost',
        'RestructuringCostsAndAssetImpairmentCharges',
        'RestructuringSettlementAndImpairmentProvisions',
    ],
    'debt_extinguishment': [
        'GainsLossesOnExtinguishmentOfDebt',
        'GainLossOnExtinguishmentOfDebt',
        'ExtinguishmentOfDebtAmount',
    ],
    'nonop_net': [
        'NonoperatingIncomeExpense',
        'OtherNonoperatingIncomeExpense',
    ],
    'interest_expense': [
        'InterestExpense',
        'InterestAndDebtExpense',
        'InterestExpenseDebt',
        'InterestExpenseNonoperating',
        'InterestExpenseRelatedParty',
        'InterestExpenseLongTermDebt',
        'InterestExpenseOther',
        'InterestExpenseBorrowings',
        'InterestCostsIncurred',
        'InterestPaidNet',
    ],
    'interest_income': [
        'InvestmentIncomeInterest',
        'InterestAndDividendIncomeOperating',
        'InterestIncomeOperating',
        'InterestIncomeOther',
        'InvestmentIncomeInterestAndDividend',
        'InvestmentIncomeNet',
    ],
    'pretax_income': [
        'IncomeLossFromContinuingOperationsBeforeIncomeTaxesExtraordinaryItemsNoncontrollingInterest',
        'IncomeLossFromContinuingOperationsBeforeIncomeTaxesMinorityInterestAndIncomeLossFromEquityMethodInvestments',
        'IncomeLossFromContinuingOperationsBeforeIncomeTaxesDomestic',
    ],
    'tax_expense': [
        'IncomeTaxExpenseBenefit',
        'CurrentIncomeTaxExpenseBenefit',
        'IncomeTaxesPaidNet',
    ],
    'net_income': [
        'NetIncomeLoss',
        'NetIncome',
        'ProfitLoss',
        'NetIncomeLossAvailableToCommonStockholdersBasic',
        'NetIncomeLossAttributableToParent',
        'IncomeLossFromContinuingOperations',
        'IncomeLossFromContinuingOperationsIncludingPortionAttributableToNoncontrollingInterest',
    ],
    'eps_basic': [
        'EarningsPerShareBasic',
        'IncomeLossFromContinuingOperationsPerBasicShare',
        'EarningsPerShareBasicAndDiluted',
    ],
    'eps_diluted': [
        'EarningsPerShareDiluted',
        'IncomeLossFromContinuingOperationsPerDilutedShare',
        'EarningsPerShareBasicAndDiluted',
    ],
    'shares_basic': [
        'WeightedAverageNumberOfSharesOutstandingBasic',
        'CommonStockSharesOutstanding',
        'CommonStockSharesIssued',
    ],
    'shares_diluted': [
        'WeightedAverageNumberOfDilutedSharesOutstanding',
        'WeightedAverageNumberOfSharesOutstandingDiluted',
        'CommonStockSharesOutstanding',
    ],
}

BS_CONCEPTS: Dict[str, List[str]] = {
    'cash': [
        'CashCashEquivalentsRestrictedCashAndRestrictedCashEquivalents',
        'CashCashEquivalentsRestrictedCashAndRestrictedCashEquivalentsIncludingDisposalGroupAndDiscontinuedOperations',
        'CashAndCashEquivalentsAtCarryingValue',
        'CashAndCashEquivalents',
        'Cash',
        'CashAndDueFromBanks',
        'CashEquivalentsAtCarryingValue',
        'RestrictedCashAndCashEquivalentsAtCarryingValue',
    ],
    'st_investments': [
        'ShortTermInvestments',
        'MarketableSecuritiesCurrent',
        'AvailableForSaleSecuritiesCurrent',
        'DebtSecuritiesAvailableForSaleCurrent',
        'TradingSecuritiesCurrent',
        'HeldToMaturitySecuritiesCurrent',
        'EquitySecuritiesFvNiCurrent',
        'InvestmentsCurrent',
    ],
    'accounts_rec': [
        'AccountsReceivableNetCurrent',
        'ReceivablesNetCurrent',
        'TradeAndOtherReceivablesNetCurrent',
        'AccountsAndNotesReceivableNet',
        'AccountsNotesAndLoansReceivableNetCurrent',
        'AccountsReceivableNet',
        'ContractWithCustomerAssetNetCurrent',
        'BilledContractReceivables',
        'AccountsReceivableGrossCurrent',
    ],
    'inventory': [
        'InventoryNet',
        'Inventories',
        'InventoryGross',
        'InventoryFinishedGoods',
        'InventoryFinishedGoodsNetOfReserves',
        'InventoryWorkInProcess',
        'InventoryRawMaterials',
        'InventoryRawMaterialsNetOfReserves',
        'InventoryAndOther',
        'SuppliesAndInventories',
        'RetailRelatedInventoryMerchandise',
        'FuelInventory',
        'InventoryNetOfAllowancesCustomerAdvancesAndProgress',
        'InventoryRealEstate',
        'AircraftMaintenanceMaterialsAndRepairs',
    ],
    'other_current_a': [
        'OtherAssetsCurrent',
        'PrepaidExpenseAndOtherAssetsCurrent',
        'PrepaidExpenseCurrent',
        'OtherCurrentAssets',
        'DeferredCostsCurrent',
        'PrepaidExpenseOtherCurrent',
        'AssetsCurrentOther',
    ],
    'total_current_a': ['AssetsCurrent'],
    'ppe_net': [
        'PropertyPlantAndEquipmentNet',
        'PropertyPlantAndEquipmentNetExcludingCapitalizedInterest',
        'PropertyAndEquipmentNet',
        'PropertyPlantAndEquipmentAndRightOfUseAssetAfterAccumulatedDepreciationAndAmortization',
        'PropertyPlantAndEquipmentAndFinanceLeaseRightOfUseAssetAfterAccumulatedDepreciationAndAmortization',
        'PropertyPlantEquipmentAndFinanceLeaseRightOfUseAssetNet',
        'PropertyPlantAndEquipmentOtherNet',
        'RealEstateInvestmentPropertyNet',
        'ElectricUtilityPlantNet',
    ],
    'goodwill': ['Goodwill', 'GoodwillGross'],
    'intangibles': [
        'IntangibleAssetsNetExcludingGoodwill',
        'FiniteLivedIntangibleAssetsNet',
        'IndefiniteLivedIntangibleAssetsExcludingGoodwill',
        'IntangibleAssetsNet',
        'OtherIntangibleAssetsNet',
        'CapitalizedComputerSoftwareNet',
    ],
    'lt_investments': [
        'MarketableSecuritiesNoncurrent',
        'AvailableForSaleSecuritiesNoncurrent',
        'OtherLongTermInvestments',
        'LongTermInvestments',
        'InvestmentsAndOtherNoncurrentAssets',
        'EquitySecuritiesWithoutReadilyDeterminableFairValueAmount',
        'EquityMethodInvestments',
    ],
    'other_noncurrent_a': [
        'OtherAssetsNoncurrent',
        'OtherNoncurrentAssets',
        'OtherLongTermAssets',
        'DeferredIncomeTaxAssetsNet',
        'OperatingLeaseRightOfUseAsset',
        'NoncurrentAssets',
    ],
    'total_assets': ['Assets'],
    'accounts_pay': [
        'AccountsPayableCurrent',
        'AccountsPayable',
        'AccountsPayableAndAccruedLiabilitiesCurrent',
        'TradeAccountsPayableCurrent',
        'AccountsPayableRelatedPartiesCurrent',
        'AccountsPayableTradeCurrent',
        'AccountsPayableCurrentAndNoncurrent',
    ],
    'accrued_liab': [
        'AccruedLiabilitiesCurrent',
        'EmployeeRelatedLiabilitiesCurrent',
        'AccruedAndOtherCurrentLiabilities',
        'AccruedExpensesAndOtherCurrentLiabilities',
        'OtherAccruedLiabilitiesCurrent',
        'AccruedExpensesCurrent',
    ],
    'other_current_l': [
        'OtherLiabilitiesCurrent',
        'OtherCurrentLiabilities',
        'AccruedIncomeTaxesCurrent',
        'TaxesPayableCurrent',
        'OtherCurrentLiabilitiesAndAccruedExpenses',
    ],
    'st_debt': [
        'ShortTermBorrowings',
        'LongTermDebtCurrent',
        'CurrentMaturitiesOfLongTermDebt',
        'DebtCurrent',
        'NotesPayableCurrent',
        'ShortTermDebtAndCurrentMaturitiesOfLongTermDebt',
        'LongTermDebtAndCapitalLeaseObligationsCurrent',
        'CommercialPaper',
        'LinesOfCreditCurrent',
    ],
    'deferred_rev_cur': [
        'DeferredRevenueCurrent',
        'ContractWithCustomerLiabilityCurrent',
        'DeferredRevenueAndCredits',
        'CustomerAdvancesCurrent',
        'DeferredIncomeCurrent',
    ],
    'total_current_l': ['LiabilitiesCurrent'],
    'lt_debt': [
        'LongTermDebtNoncurrent',
        'LongTermDebt',
        'LongTermDebtExcludingCurrentMaturities',
        'LongTermDebtAndCapitalLeaseObligations',
        'LongTermNotesPayable',
        'SeniorLongTermNotes',
        'ConvertibleDebtNoncurrent',
        'UnsecuredLongTermDebt',
        'LongTermBorrowings',
    ],
    'deferred_tax_l': [
        'DeferredIncomeTaxLiabilitiesNet',
        'DeferredTaxLiabilitiesNoncurrent',
        'DeferredTaxLiabilities',
        'DeferredIncomeTaxesAndOtherTaxLiabilitiesNoncurrent',
        'DeferredTaxAndOtherLiabilitiesNoncurrent',
        'AccruedIncomeTaxesNoncurrent',
    ],
    'other_noncurrent_l': [
        'OtherLiabilitiesNoncurrent',
        'OtherNoncurrentLiabilities',
        'OtherLongTermLiabilities',
        'LiabilitiesOtherThanLongtermDebtNoncurrent',
        'OperatingLeaseLiabilityNoncurrent',
    ],
    'total_liabilities': ['Liabilities'],
    'common_stock': [
        'CommonStockValue',
        'CommonStockValueOutstanding',
        'CommonStocksIncludingAdditionalPaidInCapital',
    ],
    'apic': [
        'AdditionalPaidInCapital',
        'AdditionalPaidInCapitalCommonStock',
    ],
    'retained_earnings': [
        'RetainedEarningsAccumulatedDeficit',
        'RetainedEarnings',
        'RetainedEarningsAppropriated',
    ],
    'treasury_stock': [
        'TreasuryStockValue',
        'TreasuryStockCommonValue',
        'TreasuryStockAtCost',
    ],
    'total_equity': [
        'StockholdersEquityIncludingPortionAttributableToNoncontrollingInterest',
        'StockholdersEquity',
        'PartnersCapital',
        'MembersEquity',
        'LimitedLiabilityCompanyLlcMembersEquityIncludingPortionAttributableToNoncontrollingInterest',
    ],
}

CF_CONCEPTS: Dict[str, List[str]] = {
    'operating_cf': [
        'NetCashProvidedByUsedInOperatingActivities',
        'NetCashProvidedByUsedInOperatingActivitiesContinuingOperations',
        'CashGeneratedFromOperations',
    ],
    'capex': [
        'PaymentsToAcquirePropertyPlantAndEquipment',
        'AcquisitionsOfPropertyPlantAndEquipment',
        'PurchasesOfPropertyAndEquipment',
        'PaymentsForCapitalImprovements',
        'PaymentsToAcquirePropertyPlantEquipmentAndOther',
        'PaymentsToAcquireProductiveAssets',
        'InvestmentsInPropertyPlantAndEquipment',
        'PaymentsToAcquireAndDevelopRealEstate',
        'CapitalExpenditureDiscontinuedOperations',
        'PaymentsToAcquireOtherPropertyPlantAndEquipment',
        'PropertyPlantAndEquipmentAdditions',
    ],
    'acquisitions': [
        'PaymentsToAcquireBusinessesNetOfCashAcquired',
        'PaymentsToAcquireBusinessesGross',
        'PaymentsToAcquireBusinessesAndInterestInAffiliates',
        'BusinessAcquisitionCostOfAcquiredEntityTransactionCosts',
    ],
    'investing_cf': [
        'NetCashProvidedByUsedInInvestingActivities',
        'NetCashProvidedByUsedInInvestingActivitiesContinuingOperations',
    ],
    'dividends': [
        'PaymentsOfDividends',
        'PaymentsOfDividendsCommonStock',
        'PaymentsOfOrdinaryDividends',
        'DividendsPaid',
        'PaymentsOfDividendsAndDividendEquivalentsOnRestrictedStockUnitsAndOtherEquityAwards',
    ],
    'repurchases': [
        'PaymentsForRepurchaseOfCommonStock',
        'TreasuryStockValueAcquiredCostMethod',
        'PaymentsForRepurchaseOfEquity',
        'PaymentsForRepurchaseOfCommonStockAndStockAwards',
        'RepurchaseOfCommonStock',
    ],
    'debt_issuance': [
        'ProceedsFromIssuanceOfLongTermDebt',
        'ProceedsFromDebtNetOfIssuanceCosts',
        'ProceedsFromIssuanceOfDebt',
        'ProceedsFromBorrowings',
        'ProceedsFromLongTermLinesOfCredit',
        'ProceedsFromIssuanceOfSeniorLongTermDebt',
        'ProceedsFromIssuanceOfUnsecuredDebt',
    ],
    'debt_repay': [
        'RepaymentsOfLongTermDebt',
        'RepaymentsOfDebt',
        'RepaymentsOfBorrowings',
        'RepaymentOfLongTermDebt',
        'RepaymentsOfSeniorDebt',
        'RepaymentsOfLinesOfCredit',
        'RepaymentsOfUnsecuredDebt',
        'RepaymentsOfNotesPayable',
    ],
    'financing_cf': [
        'NetCashProvidedByUsedInFinancingActivities',
        'NetCashProvidedByUsedInFinancingActivitiesContinuingOperations',
    ],
    'fx_effect': [
        'EffectOfExchangeRateOnCashCashEquivalentsRestrictedCashAndRestrictedCashEquivalentsIncludingDisposalGroupAndDiscontinuedOperations',
        'EffectOfExchangeRateOnCashCashEquivalentsRestrictedCashAndRestrictedCashEquivalents',
        'EffectOfExchangeRateOnCashAndCashEquivalents',
        'EffectOfExchangeRateOnCashAndCashEquivalentsContinuingOperations',
        'ExchangeRateEffectOnCashAndCashEquivalents',
        'EffectOfExchangeRateOnCash',
    ],
    'sbc': [
        'ShareBasedCompensation',
        'AllocatedShareBasedCompensationExpense',
        'ShareBasedCompensationExpense',
        'EmployeeBenefitsAndShareBasedCompensation',
        'StockBasedCompensation',
    ],
}


# =========================================================================
# QUARTERLY DATA EXTRACTION
# =========================================================================

def _get_quarterly_values(facts: Dict, concept: str,
                          fiscal_year: int,
                          taxonomy: str = 'us-gaap') -> List[Dict]:
    """Extract 10-Q values for a concept in a given fiscal year.

    Returns list of fact entries sorted by period end date (newest first).
    """
    try:
        concept_data = facts['facts'][taxonomy][concept]
        units = concept_data.get('units', {})
        for unit_key in ('USD', 'shares', 'pure'):
            if unit_key not in units:
                continue
            quarterly = [
                v for v in units[unit_key]
                if v.get('form') == '10-Q' and v.get('fy') == fiscal_year
                and 'end' in v and 'val' in v and v.get('fp')
                and v.get('fp') != 'FY'
            ]
            if not quarterly:
                continue
            # Deduplicate: keep latest filing per (fp, start, end)
            best = {}
            for v in quarterly:
                key = (v.get('fp'), v.get('start', ''), v['end'])
                if key not in best or v.get('filed', '') > best[key].get('filed', ''):
                    best[key] = v
            return sorted(best.values(), key=lambda x: x['end'], reverse=True)
        return []
    except (KeyError, TypeError):
        return []


def _get_quarterly_metric_for_fp(facts: Dict, concepts: List[str],
                                 fiscal_year: int, target_fp: str,
                                 period_type: str = 'ytd',
                                 taxonomy: str = 'us-gaap') -> Optional[float]:
    """Get a single quarterly metric value for a specific fiscal period.

    Args:
        period_type: 'ytd' picks the longest-duration value (cumulative).
                     'instant' picks point-in-time values (balance sheet).
    """
    for concept in concepts:
        values = _get_quarterly_values(facts, concept, fiscal_year, taxonomy)
        if not values:
            continue
        # Filter to the target quarter
        fp_vals = [v for v in values if v.get('fp') == target_fp]
        if not fp_vals:
            continue

        if period_type == 'instant':
            # BS: just return the value (point-in-time)
            return float(fp_vals[0]['val'])

        # IS/CF: pick the value with the longest period (YTD cumulative)
        best_val = None
        best_duration = -1
        for v in fp_vals:
            start = v.get('start', '')
            end = v.get('end', '')
            if start and end:
                try:
                    d = (datetime.strptime(end, '%Y-%m-%d')
                         - datetime.strptime(start, '%Y-%m-%d')).days
                except ValueError:
                    d = 0
            else:
                d = 0
            if d > best_duration:
                best_duration = d
                best_val = float(v['val'])
        if best_val is not None:
            return best_val
    return None


def _identify_latest_quarter(facts: Dict,
                             latest_annual_year: int) -> Optional[Dict]:
    """Find the most recent 10-Q quarter filed AFTER the latest annual year.

    Returns {'fy': int, 'fp': str} or None.
    Skips Q4 (covered by 10-K).
    """
    probe_concepts = [
        'Assets', 'Revenues', 'NetIncomeLoss',
        'RevenueFromContractWithCustomerExcludingAssessedTax',
    ]
    best = None  # (fy, fp, end_date)
    for concept in probe_concepts:
        try:
            concept_data = facts['facts']['us-gaap'][concept]
            units = concept_data.get('units', {})
            for unit_key in ('USD', 'shares', 'pure'):
                if unit_key not in units:
                    continue
                for v in units[unit_key]:
                    if v.get('form') != '10-Q':
                        continue
                    fp = v.get('fp', '')
                    fy = v.get('fy')
                    if not fy or fp not in ('Q1', 'Q2', 'Q3'):
                        continue
                    # Must be after the latest annual year
                    if fy <= latest_annual_year:
                        continue
                    end = v.get('end', '')
                    if best is None or end > best[2]:
                        best = (fy, fp, end)
        except (KeyError, TypeError):
            continue
    if best is None:
        return None
    return {'fy': best[0], 'fp': best[1]}


def _recompute_derived(fd: Dict, yr) -> None:
    """Recompute GP, EBITDA, FCF, other_income for a given year key."""
    inc = fd['income_statement']
    cf = fd['cash_flow']

    # Gross Profit
    rev = inc['revenue'].get(yr)
    cogs_val = inc['cogs'].get(yr)
    gp = inc['gross_profit'].get(yr)
    if gp is None and rev is not None and cogs_val is not None:
        inc['gross_profit'][yr] = rev - cogs_val
    elif cogs_val is None and rev is not None and gp is not None:
        inc['cogs'][yr] = rev - gp

    # EBITDA = OI + D&A + Amort + Transform + |DebtExt|
    oi = inc['operating_income'].get(yr)
    da_val = inc['da'].get(yr)
    if oi is not None and da_val is not None:
        am = inc.get('amortization', {}).get(yr) or 0
        tc = inc.get('transformation_costs', {}).get(yr) or 0
        de = abs(inc.get('debt_extinguishment', {}).get(yr) or 0)
        inc['ebitda'][yr] = oi + da_val + am + tc + de
    elif oi is not None:
        inc['ebitda'][yr] = None

    # Other income plug: PreTax - EBIT - IntInc + IntExp
    pt = inc.get('pretax_income', {}).get(yr)
    ebit = inc.get('operating_income', {}).get(yr)
    if pt is not None and ebit is not None:
        ii = inc.get('interest_income', {}).get(yr) or 0
        ie = inc.get('interest_expense', {}).get(yr) or 0
        inc['other_income'][yr] = pt - ebit - ii + ie

    # FCF = Operating CF - |Capex|
    opcf = cf['operating_cf'].get(yr)
    capex_val = cf['capex'].get(yr)
    if opcf is not None and capex_val is not None:
        cf['fcf'][yr] = opcf - abs(capex_val)


# Quarter ordering for comparison
_Q_ORDER = {'Q1': 1, 'Q2': 2, 'Q3': 3}


def compute_ltm(facts: Dict, financial_data: Dict,
                ticker: Optional[str] = None) -> Optional[int]:
    """Compute quarterly cumulative and LTM annualized financials from 10-Q data.

    Produces TWO sets of data when quarterly data is available:

    1. **Q cumulative** -- YTD values from the latest 10-Q, stored under
       ``ltm_year`` in the normal data dicts.  Column label: ``'Q3 FY2025'``.

    2. **Annualized** -- Trailing 12 months:
       ``Annual(base_year) + Q_cum(current) - Q_cum(prior)``
       Column label: ``'Ann.Q3 FY2025'``.

    Returns the ltm_year if successful, or None.
    """
    latest_annual_year = financial_data['years'][0]  # newest first
    q_info = _identify_latest_quarter(facts, latest_annual_year)
    if q_info is None:
        return None

    q_fy = q_info['fy']
    q_fp = q_info['fp']  # 'Q1', 'Q2', 'Q3'

    print(f"  Found {q_fp} FY{q_fy} quarterly data, computing LTM...")

    years = financial_data['years']
    ltm_year = q_fy
    base_year = latest_annual_year  # most recent full fiscal year

    # Add ltm_year to years list (newest first)
    if ltm_year not in years:
        years.insert(0, ltm_year)
        financial_data['years'] = years

    ann_is: Dict[str, Optional[float]] = {}
    ann_cf: Dict[str, Optional[float]] = {}
    ann_bs: Dict[str, Optional[float]] = {}

    # --- IS metrics ---
    for key, concepts in IS_CONCEPTS.items():
        if key not in financial_data['income_statement']:
            continue
        # Get YTD cumulative for current and prior year same quarter
        q_cur = _get_quarterly_metric_for_fp(facts, concepts, q_fy, q_fp, 'ytd')
        q_pri = _get_quarterly_metric_for_fp(facts, concepts, q_fy - 1, q_fp, 'ytd')
        annual_val = financial_data['income_statement'][key].get(base_year)

        # Store Q cumulative in normal data dict
        financial_data['income_statement'][key][ltm_year] = q_cur

        # Annualized = Annual + Q_cur - Q_pri
        if annual_val is not None and q_cur is not None and q_pri is not None:
            ann_is[key] = annual_val + q_cur - q_pri
        else:
            ann_is[key] = q_cur  # fallback: just use Q cumulative

    # --- CF metrics ---
    for key, concepts in CF_CONCEPTS.items():
        if key not in financial_data['cash_flow']:
            continue
        q_cur = _get_quarterly_metric_for_fp(facts, concepts, q_fy, q_fp, 'ytd')
        q_pri = _get_quarterly_metric_for_fp(facts, concepts, q_fy - 1, q_fp, 'ytd')
        annual_val = financial_data['cash_flow'][key].get(base_year)

        financial_data['cash_flow'][key][ltm_year] = q_cur

        if annual_val is not None and q_cur is not None and q_pri is not None:
            ann_cf[key] = annual_val + q_cur - q_pri
        else:
            ann_cf[key] = q_cur

    # --- BS metrics (point-in-time: same for Q and Ann) ---
    for key, concepts in BS_CONCEPTS.items():
        if key not in financial_data['balance_sheet']:
            continue
        val = _get_quarterly_metric_for_fp(facts, concepts, q_fy, q_fp, 'instant')
        financial_data['balance_sheet'][key][ltm_year] = val
        ann_bs[key] = val

    # Total Liabilities plug for BS
    ta = financial_data['balance_sheet']['total_assets'].get(ltm_year)
    te = financial_data['balance_sheet']['total_equity'].get(ltm_year)
    if ta is not None and te is not None:
        financial_data['balance_sheet']['total_liabilities'][ltm_year] = ta - te
        ann_bs['total_liabilities'] = ta - te

    # Recompute derived metrics for Q cumulative
    _recompute_derived(financial_data, ltm_year)

    # --- Annualized derived metrics ---
    # GP
    ann_rev = ann_is.get('revenue')
    ann_cogs = ann_is.get('cogs')
    if ann_is.get('gross_profit') is None and ann_rev and ann_cogs:
        ann_is['gross_profit'] = ann_rev - ann_cogs
    elif ann_is.get('cogs') is None and ann_rev and ann_is.get('gross_profit'):
        ann_is['cogs'] = ann_rev - ann_is['gross_profit']

    # EBITDA
    ann_oi = ann_is.get('operating_income')
    ann_da = ann_is.get('da')
    if ann_oi is not None and ann_da is not None:
        ann_is['ebitda'] = (ann_oi + ann_da
                            + (ann_is.get('amortization') or 0)
                            + (ann_is.get('transformation_costs') or 0)
                            + abs(ann_is.get('debt_extinguishment') or 0))
    else:
        ann_is['ebitda'] = None

    # Other income plug
    ann_pt = ann_is.get('pretax_income')
    if ann_pt is not None and ann_oi is not None:
        ann_ii = ann_is.get('interest_income') or 0
        ann_ie = ann_is.get('interest_expense') or 0
        ann_is['other_income'] = ann_pt - ann_oi - ann_ii + ann_ie

    # FCF
    ann_opcf = ann_cf.get('operating_cf')
    ann_capex = ann_cf.get('capex')
    if ann_opcf is not None and ann_capex is not None:
        ann_cf['fcf'] = ann_opcf - abs(ann_capex)
    else:
        ann_cf['fcf'] = None

    # Store annualized data and metadata
    q_short = q_fp  # 'Q1', 'Q2', 'Q3'
    financial_data['annualized'] = {
        'label': f'Ann.{q_short} FY{q_fy}',
        'year': q_fy,
        'income_statement': ann_is,
        'balance_sheet': ann_bs,
        'cash_flow': ann_cf,
    }
    financial_data['ltm_info'] = {
        'ltm_year': ltm_year,
        'quarter': q_short,
        'base_year': base_year,
        'q_label': f'{q_short} FY{q_fy}',
        'ann_label': f'Ann.{q_short} FY{q_fy}',
    }

    return ltm_year


def extract_financial_data(facts: Dict, years: List[int],
                           ticker: Optional[str] = None) -> Dict:
    """Extract all 3-statement financial data for the given fiscal years.

    Each metric tries an ordered list of XBRL concept names (most preferred first).
    The first concept that returns data for ALL requested years is used; if none
    covers every year, the function fills as many years as possible by combining
    results across concepts, so partial data is still captured.

    *ticker* (optional) enables yfinance fallback for shares outstanding when
    XBRL data is completely missing (e.g. BRK-B, V).
    """
    print(f"  Extracting financial statements for FY{years[-1]}-FY{years[0]}...")

    def g(concepts, tax='us-gaap'):
        return _get_metric(facts, concepts, years, tax)

    # =========================================================================
    # INCOME STATEMENT
    # =========================================================================

    revenue = g(IS_CONCEPTS['revenue'])
    cogs = g(IS_CONCEPTS['cogs'])
    gross_profit = g(IS_CONCEPTS['gross_profit'])
    # Derive COGS or Gross Profit from the other when one is missing:
    # GP = Revenue - COGS, so COGS = Revenue - GP.
    for y in years:
        if gross_profit[y] is None and revenue[y] and cogs[y]:
            gross_profit[y] = revenue[y] - cogs[y]
        elif cogs[y] is None and revenue[y] and gross_profit[y] is not None:
            cogs[y] = revenue[y] - gross_profit[y]

    rd_expense = g(IS_CONCEPTS['rd_expense'])

    # SGA: try combined concept first; if absent, sum G&A + Sales & Marketing
    sga = g(IS_CONCEPTS['sga_expense'])
    if all(v is None for v in sga.values()):
        ga = g(IS_CONCEPTS['sga_ga'])
        sm = g(IS_CONCEPTS['sga_sm'])
        sga = {y: ((ga.get(y) or 0) + (sm.get(y) or 0))
                   if ga.get(y) is not None or sm.get(y) is not None
                   else None
               for y in years}

    operating_income = g(IS_CONCEPTS['operating_income'])
    da = g(IS_CONCEPTS['da'])
    amortization = g(IS_CONCEPTS['amortization'])
    transformation_costs = g(IS_CONCEPTS['transformation_costs'])
    debt_extinguishment = g(IS_CONCEPTS['debt_extinguishment'])

    # EBITDA derived from EBIT + D&A + add-backs
    ebitda: Dict[int, Optional[float]] = {}
    for y in years:
        if operating_income[y] is not None and da[y] is not None:
            ebitda[y] = (operating_income[y] + da[y]
                         + (amortization[y] or 0)
                         + (transformation_costs[y] or 0)
                         + abs(debt_extinguishment[y] or 0))
        else:
            ebitda[y] = None

    interest_expense = g(IS_CONCEPTS['interest_expense'])
    interest_income = g(IS_CONCEPTS['interest_income'])

    # Other income / (expense) — non-operating items not captured by
    # interest income or interest expense (e.g. investment gains/losses,
    # FX gains/losses, equity method income).  Derived as:
    #   PreTax - EBIT - InterestIncome + InterestExpense
    # so that EBIT + IntInc - IntExp + OtherIncome = PreTax exactly.
    # Computed after pretax_income is extracted (see below).

    nonop_net = g(IS_CONCEPTS['nonop_net'])
    pretax_income = g(IS_CONCEPTS['pretax_income'])

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

    tax_expense = g(IS_CONCEPTS['tax_expense'])
    net_income = g(IS_CONCEPTS['net_income'])
    eps_basic = g(IS_CONCEPTS['eps_basic'])
    eps_diluted = g(IS_CONCEPTS['eps_diluted'])
    shares_basic = g(IS_CONCEPTS['shares_basic'])
    shares_diluted = g(IS_CONCEPTS['shares_diluted'])

    # Fallback: try dei (Document/Entity Information) taxonomy for entity-level
    # shares when us-gaap weighted-average concepts are not available (BRK-B, V).
    if all(v is None or v == 0 for v in shares_diluted.values()):
        dei_shares = _get_metric(facts, ['EntityCommonStockSharesOutstanding'], years,
                                 taxonomy='dei')
        for y in years:
            if (shares_diluted[y] is None or shares_diluted[y] == 0) and dei_shares.get(y):
                shares_diluted[y] = dei_shares[y]
    if all(v is None or v == 0 for v in shares_basic.values()):
        dei_shares = _get_metric(facts, ['EntityCommonStockSharesOutstanding'], years,
                                 taxonomy='dei')
        for y in years:
            if (shares_basic[y] is None or shares_basic[y] == 0) and dei_shares.get(y):
                shares_basic[y] = dei_shares[y]

    # Normalize shares: some filers (e.g. MCD FY2023+) report weighted-average
    # shares in millions instead of raw units.  We detect per-year anomalies by
    # looking for values < 10,000 (no S&P 500 company has fewer than 10K diluted
    # shares).  If a mix of large (>1M) and small (<10K) values exists (unit
    # change across years), the small values are multiplied by 1e6.
    def _normalize_shares(d: Dict[int, Optional[float]]) -> None:
        filled = {y: v for y, v in d.items() if v is not None and v > 0}
        if not filled:
            return
        for y in d:
            if d[y] is not None and d[y] > 0 and d[y] < 10_000:
                # Value is suspiciously small — likely in millions.
                # Scale up whether or not other years are large (consistent
                # millions-reporting or mixed-unit reporting both need fixing).
                d[y] = d[y] * 1_000_000
    _normalize_shares(shares_basic)
    _normalize_shares(shares_diluted)

    # Helper: check if shares data is effectively missing (all None or all zero).
    def _shares_missing(d: Dict[int, Optional[float]]) -> bool:
        return all(v is None or v == 0 for v in d.values())

    # yfinance fallback for shares when XBRL has no data (BRK-B, V, etc.)
    if _shares_missing(shares_diluted):
        # First try: back-calculate from NI / EPS when both are tagged.
        _eps = _get_metric(facts, [
            'EarningsPerShareDiluted',
            'IncomeLossFromContinuingOperationsPerDilutedShare',
            'EarningsPerShareBasicAndDiluted',
        ], years)
        _ni = _get_metric(facts, [
            'NetIncomeLoss', 'ProfitLoss',
            'NetIncomeLossAvailableToCommonStockholdersBasic',
        ], years)
        for y in years:
            if (_eps.get(y) and _ni.get(y) and _eps[y] != 0):
                shares_diluted[y] = _ni[y] / _eps[y]
                if shares_basic[y] is None or shares_basic[y] == 0:
                    shares_basic[y] = shares_diluted[y]

    if _shares_missing(shares_diluted) and _HAS_YFINANCE and ticker:
        # Second try: use yfinance sharesOutstanding as a constant for all years.
        try:
            stock = yf.Ticker(ticker)
            yf_shares = stock.info.get('sharesOutstanding')
            if yf_shares and yf_shares > 0:
                print(f"  Note: using yfinance shares ({yf_shares/1e6:,.0f}M) "
                      f"- XBRL shares not tagged for {ticker}")
                for y in years:
                    shares_diluted[y] = float(yf_shares)
                    if shares_basic[y] is None:
                        shares_basic[y] = float(yf_shares)
        except Exception:
            pass

    # Forward-fill shares: if some years have data but others are 0/None,
    # carry forward the most recent non-zero value.  Shares outstanding change
    # slowly year-over-year, so a prior year's count is a reasonable proxy.
    def _ffill_shares(d: Dict[int, Optional[float]]) -> None:
        sorted_years = sorted(d.keys())
        last_good = None
        for y in sorted_years:
            if d[y] is not None and d[y] > 0:
                last_good = d[y]
            elif last_good is not None:
                d[y] = last_good
    _ffill_shares(shares_basic)
    _ffill_shares(shares_diluted)

    # =========================================================================
    # BALANCE SHEET – ASSETS
    # =========================================================================

    cash = g(BS_CONCEPTS['cash'])
    st_investments = g(BS_CONCEPTS['st_investments'])
    accounts_rec = g(BS_CONCEPTS['accounts_rec'])
    inventory = g(BS_CONCEPTS['inventory'])
    other_current_a = g(BS_CONCEPTS['other_current_a'])
    total_current_a = g(BS_CONCEPTS['total_current_a'])
    ppe_net = g(BS_CONCEPTS['ppe_net'])
    goodwill = g(BS_CONCEPTS['goodwill'])
    intangibles = g(BS_CONCEPTS['intangibles'])
    lt_investments = g(BS_CONCEPTS['lt_investments'])
    other_noncurrent_a = g(BS_CONCEPTS['other_noncurrent_a'])
    total_assets = g(BS_CONCEPTS['total_assets'])

    # =========================================================================
    # BALANCE SHEET – LIABILITIES
    # =========================================================================

    accounts_pay = g(BS_CONCEPTS['accounts_pay'])
    accrued_liab = g(BS_CONCEPTS['accrued_liab'])
    other_current_l = g(BS_CONCEPTS['other_current_l'])
    st_debt = g(BS_CONCEPTS['st_debt'])
    deferred_rev_cur = g(BS_CONCEPTS['deferred_rev_cur'])
    total_current_l = g(BS_CONCEPTS['total_current_l'])
    lt_debt = g(BS_CONCEPTS['lt_debt'])
    deferred_tax_l = g(BS_CONCEPTS['deferred_tax_l'])
    other_noncurrent_l = g(BS_CONCEPTS['other_noncurrent_l'])
    total_liabilities = g(BS_CONCEPTS['total_liabilities'])

    # =========================================================================
    # BALANCE SHEET – EQUITY
    # =========================================================================

    common_stock = g(BS_CONCEPTS['common_stock'])
    apic = g(BS_CONCEPTS['apic'])
    retained_earnings = g(BS_CONCEPTS['retained_earnings'])
    treasury_stock = g(BS_CONCEPTS['treasury_stock'])
    total_equity = g(BS_CONCEPTS['total_equity'])

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

    operating_cf = g(CF_CONCEPTS['operating_cf'])
    capex = g(CF_CONCEPTS['capex'])
    acquisitions = g(CF_CONCEPTS['acquisitions'])
    investing_cf = g(CF_CONCEPTS['investing_cf'])
    dividends = g(CF_CONCEPTS['dividends'])
    repurchases = g(CF_CONCEPTS['repurchases'])
    debt_issuance = g(CF_CONCEPTS['debt_issuance'])
    debt_repay = g(CF_CONCEPTS['debt_repay'])
    financing_cf = g(CF_CONCEPTS['financing_cf'])
    fx_effect = g(CF_CONCEPTS['fx_effect'])
    sbc = g(CF_CONCEPTS['sbc'])

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
            'amortization':     amortization,
            'transformation_costs': transformation_costs,
            'debt_extinguishment':  debt_extinguishment,
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
