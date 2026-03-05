"""
XBRL Tag Investigation
======================
For each problematic company, fetch XBRL data and find what tags they
actually use for the financial items our model is missing.
"""

import json
import sys
from sec_fetcher import search_company, get_company_facts, get_fiscal_years, _get_annual_values

# Companies with specific issues to investigate
INVESTIGATE = {
    'MCD':   ['shares', 'cogs', 'accounts_rec'],
    'BRK-B': ['revenue', 'shares', 'operating_income'],
    'V':     ['revenue', 'shares', 'operating_income'],
    'JPM':   ['revenue', 'cogs', 'operating_income'],
    'NVDA':  ['revenue', 'cogs', 'operating_income', 'da'],
    'NEE':   ['cogs', 'capex', 'interest_expense', 'accounts_pay'],
    'GE':    ['cogs', 'interest_expense', 'ppe'],
    'AAPL':  ['interest_expense'],
    'AMZN':  ['cogs', 'operating_income'],
    'TSLA':  ['cogs', 'operating_income', 'da'],
    'TMO':   ['cogs', 'interest_expense'],
    'LIN':   ['cogs'],
    'COST':  ['cogs'],
    'INTU':  ['cogs'],
    'CRM':   ['accounts_pay'],
    'DHR':   ['interest_expense', 'accounts_pay'],
    'NOW':   ['interest_expense'],
    'AMD':   ['tax_expense'],
    'AMGN':  ['interest_expense'],
    'TXN':   ['capex'],
}

# Tags to search for by category
SEARCH_TAGS = {
    'revenue': [
        'Revenues', 'RevenueFromContractWithCustomerExcludingAssessedTax',
        'RevenueFromContractWithCustomerIncludingAssessedTax',
        'SalesRevenueNet', 'SalesRevenueGoodsNet', 'SalesRevenueServicesNet',
        'RevenuesNetOfInterestExpense', 'InterestAndFeeIncomeLoansAndLeases',
        'RevenueFromContractWithCustomer', 'Revenue',
        'TotalRevenuesAndOtherIncome', 'RegulatedAndUnregulatedOperatingRevenue',
        'ElectricUtilityRevenue', 'InsurancePremiumsRevenueNet',
        'InterestIncomeExpenseNet', 'TotalRevenues',
        'InsuranceServicesRevenue', 'PremiumsEarnedNet',
    ],
    'cogs': [
        'CostOfRevenue', 'CostOfGoodsAndServicesSold', 'CostOfGoodsSold',
        'CostOfServices', 'CostOfGoodsAndServicesExcludingDepreciationDepletionAndAmortization',
        'CostOfRevenueExcludingDepreciationAndAmortization',
        'CostOfGoodsSoldExcludingDepreciationDepletionAndAmortization',
        'BenefitsLossesAndExpenses', 'PolicyholderBenefitsAndClaimsIncurredNet',
        'CostOfGoodsSoldDirectMaterials', 'CostsAndExpenses',
        'OperatingCostsAndExpenses', 'DirectOperatingCosts',
        'FoodAndBeverageCostOfRevenue', 'MediaCostOfRevenue',
        'OccupancyCosts', 'DirectCostsOfLeasedAndRentedPropertyOrEquipment',
        'FranchisorCosts', 'FranchisorRevenueCosts',
    ],
    'operating_income': [
        'OperatingIncomeLoss',
        'IncomeLossFromContinuingOperationsBeforeIncomeTaxesExtraordinaryItemsNoncontrollingInterest',
        'IncomeLossFromContinuingOperationsBeforeIncomeTaxesMinorityInterestAndIncomeLossFromEquityMethodInvestments',
    ],
    'shares': [
        'WeightedAverageNumberOfDilutedSharesOutstanding',
        'WeightedAverageNumberOfSharesOutstandingDiluted',
        'CommonStockSharesOutstanding', 'CommonStockSharesIssued',
        'WeightedAverageNumberOfSharesOutstandingBasic',
        'EntityCommonStockSharesOutstanding',
    ],
    'interest_expense': [
        'InterestExpense', 'InterestAndDebtExpense', 'InterestExpenseDebt',
        'InterestExpenseRelatedParty', 'FinanceLeaseInterestExpense',
        'InterestExpenseLongTermDebt', 'InterestExpenseOther',
        'InterestPaidNet', 'InterestCostsIncurred',
        'InterestExpenseDebtExcludingAmortization',
        'InterestIncomeExpenseNonoperatingNet',
        'InterestExpenseBorrowings',
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
        'NoncashOrPartNoncashAcquisitionFixedAssetsAcquired1',
        'CapitalExpendituresIncurredButNotYetPaid',
    ],
    'ppe': [
        'PropertyPlantAndEquipmentNet',
        'PropertyPlantAndEquipmentNetExcludingCapitalizedInterest',
        'PropertyAndEquipmentNet',
        'PropertyPlantAndEquipmentAndRightOfUseAssetAfterAccumulatedDepreciationAndAmortization',
        'PropertyPlantAndEquipmentAndFinanceLeaseRightOfUseAssetAfterAccumulatedDepreciationAndAmortization',
        'PropertyPlantEquipmentAndFinanceLeaseRightOfUseAssetNet',
        'PropertyPlantAndEquipmentOtherNet',
        'RealEstateAndAccumulatedDepreciationCarryingAmountOfBuildingsAndImprovements',
    ],
    'da': [
        'DepreciationDepletionAndAmortization',
        'DepreciationAndAmortization',
        'Depreciation',
        'AmortizationOfIntangibleAssets',
        'DepreciationAmortizationAndAccretionNet',
        'OtherDepreciationAndAmortization',
        'DepreciationAndAmortizationExcludingDisposals',
    ],
    'accounts_rec': [
        'AccountsReceivableNetCurrent',
        'ReceivablesNetCurrent',
        'TradeAndOtherReceivablesNetCurrent',
        'AccountsAndNotesReceivableNet',
        'AccountsReceivableNet',
        'ContractWithCustomerAssetNetCurrent',
        'BilledContractReceivables',
        'AccountsReceivableGrossCurrent',
    ],
    'accounts_pay': [
        'AccountsPayableCurrent',
        'AccountsPayable',
        'AccountsPayableAndAccruedLiabilitiesCurrent',
        'TradeAccountsPayableCurrent',
        'AccountsPayableRelatedPartiesCurrent',
        'AccountsPayableTradeCurrent',
    ],
    'tax_expense': [
        'IncomeTaxExpenseBenefit',
        'CurrentIncomeTaxExpenseBenefit',
        'IncomeTaxesPaidNet',
        'IncomeTaxExpenseBenefitContinuingOperations',
    ],
}


def scan_all_concepts(facts, years, category_tags):
    """Try all tags in a category and report which ones have data."""
    results = []
    us_gaap = facts.get('facts', {}).get('us-gaap', {})

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
                    val_str = f"${latest/1e6:,.0f}M" if latest and abs(latest) > 1e3 else str(latest)
                    in_model = tag in SEARCH_TAGS[issue][:len(SEARCH_TAGS[issue])]
                    print(f"    {tag}: {val_str} (FY{years[0]})")
            else:
                print(f"  No known tags found data!")

        # Broad search for related concepts
        search_keywords = {
            'revenue': ['Revenue', 'Sales', 'Income'],
            'cogs': ['Cost', 'Expense'],
            'operating_income': ['OperatingIncome', 'OperatingLoss'],
            'shares': ['Shares', 'Stock'],
            'interest_expense': ['Interest'],
            'capex': ['PropertyPlant', 'CapitalExpenditure', 'Payments'],
            'ppe': ['PropertyPlant', 'Property'],
            'da': ['Depreciation', 'Amortization'],
            'accounts_rec': ['Receivable'],
            'accounts_pay': ['Payable'],
            'tax_expense': ['IncomeTax'],
        }

        if issue in search_keywords:
            print(f"  Broad concept search:")
            for kw in search_keywords[issue]:
                broad_results = scan_all_available_concepts(facts, years, kw)
                for tag, data in broad_results:
                    latest = data.get(years[0])
                    if latest is None:
                        continue
                    val_str = f"${latest/1e6:,.1f}M" if abs(latest) > 1e3 else f"{latest:,.2f}"
                    # Only show significant values (> $1M or shares > 1000)
                    if abs(latest) > 1e6 or (issue == 'shares' and latest > 1000):
                        print(f"    {tag}: {val_str} (FY{years[0]})")


def main():
    # Investigate just the most problematic companies first
    priority = ['MCD', 'BRK-B', 'V', 'AAPL', 'NEE', 'GE', 'TMO', 'AMZN', 'TSLA']

    if len(sys.argv) > 1:
        priority = sys.argv[1:]

    for ticker in priority:
        if ticker in INVESTIGATE:
            investigate_company(ticker, INVESTIGATE[ticker])
        else:
            investigate_company(ticker, list(SEARCH_TAGS.keys()))
        import time
        time.sleep(0.3)


if __name__ == '__main__':
    main()
