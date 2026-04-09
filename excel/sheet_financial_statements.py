"""Financial Statements sheet builder."""

from typing import Dict
from .shared import *

def _write_financial_statements(ws, company_info: Dict, fd: Dict):
    years_desc = fd['years']                          # newest first
    years = list(reversed(years_desc))               # oldest first for display
    n    = len(years)
    inc  = fd['income_statement']
    bs   = fd['balance_sheet']
    cf   = fd['cash_flow']

    # ── Annualized (LTM) column setup ─────────────────────────────────
    has_ann = 'annualized' in fd and 'ltm_info' in fd
    _ltm_info = fd.get('ltm_info', {})
    ANN_YEAR = None
    if has_ann:
        ANN_YEAR = years[-1] + 0.5          # sentinel key for annualized col
        _ann_src = fd['annualized']
        _ann_is = _ann_src.get('income_statement', {})
        _ann_bs = _ann_src.get('balance_sheet', {})
        _ann_cf = _ann_src.get('cash_flow', {})
        for key in inc:
            if isinstance(inc[key], dict):
                inc[key][ANN_YEAR] = _ann_is.get(key)
        for key in bs:
            if isinstance(bs[key], dict):
                bs[key][ANN_YEAR] = _ann_bs.get(key)
        for key in cf:
            if isinstance(cf[key], dict):
                cf[key][ANN_YEAR] = _ann_cf.get(key)
        years.append(ANN_YEAR)
        n = len(years)

    # ── Projection setup ────────────────────────────────────────────────
    latest_yr      = max(yr for yr in years if isinstance(yr, int))
    proj_years     = [latest_yr + i for i in range(1, 6)]
    n_proj         = 5
    total_cols     = 1 + n + n_proj          # 10 (label + 4 hist + 5 proj)
    proj_start_col = 2 + n                   # column 6 (F)

    def _cl(i: int) -> str:
        """Excel column letter for year index i (0 = oldest year = column B)."""
        return get_column_letter(2 + i)

    def _pcl(j: int) -> str:
        """Excel column letter for projection year index j (0 = first proj year)."""
        return get_column_letter(proj_start_col + j)

    def _fw(row_num: int, col_idx: int, formula: str,
            fmt: str = FMT_DOLLAR, bold: bool = False, fill: str = None):
        """Write an Excel formula string to the data cell at (row_num, year col_idx)."""
        c = ws.cell(row=row_num, column=2 + col_idx, value=formula)
        _style(c,
               fill_hex=fill or (LIGHT_GRAY if col_idx % 2 == 0 else WHITE),
               bold=bold, h_align='right', number_format=fmt)
        if bold:
            c.border = BOT_MED
        return c

    def _pfw(row_num: int, proj_idx: int, formula: str,
             fmt: str = FMT_DOLLAR, bold: bool = False, fill: str = None):
        """Write an Excel formula into a projection column cell."""
        c = ws.cell(row=row_num, column=proj_start_col + proj_idx, value=formula)
        _style(c,
               fill_hex=fill or (XLIGHT_BLUE if proj_idx % 2 == 0 else WHITE),
               bold=bold, h_align='right', number_format=fmt)
        if bold:
            c.border = BOT_MED
        return c

    def _lbl(row_num: int, text: str, bold: bool = False,
             fill: str = None, ind: int = 0):
        c = ws.cell(row=row_num, column=1, value='  ' * ind + text)
        _style(c, fill_hex=fill, bold=bold)
        return c

    DARK_GREEN = '006400'

    def _fix_ann_headers(hdr_row):
        """Overwrite Q cumulative and Annualized column headers."""
        if not has_ann:
            return
        ltm_yr = _ltm_info['ltm_year']
        ltm_idx = years.index(ltm_yr)
        c = ws.cell(row=hdr_row, column=2 + ltm_idx,
                    value=_ltm_info['q_label'])
        _style(c, fill_hex=DARK_BLUE, bold=True, font_color=WHITE,
               h_align='center')
        c.border = THIN_BOX
        ann_idx = years.index(ANN_YEAR)
        c = ws.cell(row=hdr_row, column=2 + ann_idx,
                    value=_ltm_info['ann_label'])
        _style(c, fill_hex=DARK_GREEN, bold=True, font_color=WHITE,
               h_align='center')
        c.border = THIN_BOX

    # Column setup — historical (+ optional Q/Ann) + 5 projection columns
    col_widths = {1: 40}
    for ci in range(2, 2 + n + n_proj):
        col_widths[ci] = 16
    _set_col_widths(ws, col_widths)
    r = 1

    # Title
    title_cell = ws.cell(
        row=r, column=1,
        value=f"{company_info['name']}  ({company_info['ticker']})"
              f" - Financial Statements  ($ in millions)")
    _style(title_cell, fill_hex=DARK_BLUE, bold=True, font_color=WHITE)
    ws.merge_cells(start_row=r, start_column=1, end_row=r, end_column=total_cols)
    ws.row_dimensions[r].height = 18
    r += 1

    # Subtitle (merge up to second-to-last column, leave last proj col for dropdown)
    sub = ws.cell(
        row=r, column=1,
        value="Source: SEC EDGAR XBRL  |  Shaded rows = formula-derived  |  Amounts in $USD Millions")
    _style(sub, fill_hex=XLIGHT_BLUE, italic=True)
    ws.merge_cells(start_row=r, start_column=1, end_row=r,
                   end_column=total_cols - 1)

    # ── Scenario dropdown cell (subtitle row, last projection column) ────
    dropdown_col = proj_start_col + n_proj - 1   # column J (10)
    dropdown_ref = f'${get_column_letter(dropdown_col)}${r}'
    dd_cell = ws.cell(row=r, column=dropdown_col, value='Base Case')
    _style(dd_cell, fill_hex=YELLOW, bold=True, h_align='center')
    dd_cell.border = THIN_BOX
    dv = DataValidation(
        type='list',
        formula1='"Best Case,Base Case,Weak Case"',
        allow_blank=False,
    )
    dv.error = 'Please select Best Case, Base Case, or Weak Case'
    dv.errorTitle = 'Invalid Scenario'
    ws.add_data_validation(dv)
    dv.add(dd_cell)
    r += 2

    # =========================================================================
    # INCOME STATEMENT
    # Derived items (Gross Profit, EBITDA, EBIT, Pre-tax, Net Income, margins)
    # are Excel formulas so they update when raw inputs are edited.
    # Projection columns (5 years) are driven by the scenario dropdown.
    #
    # Layout:  Revenue - COGS = Gross Profit
    #          GP - R&D - SGA - Other_Opex = EBITDA   <- Other_Opex is a balancing plug
    #          EBITDA - D&A = EBIT
    #          EBIT - IntExp + IntInc = Pre-tax
    #          Pre-tax - Tax = Net Income
    # =========================================================================
    _write_section_header(ws, r, 'INCOME STATEMENT', cols=total_cols);  r += 1

    # Column headers — historical (DARK_BLUE) + projection (MED_BLUE)
    _write_col_headers(ws, r, list(range(2, 2 + n)), years, start_col=2)
    _fix_ann_headers(r)
    for j, py in enumerate(proj_years):
        col = proj_start_col + j
        c = ws.cell(row=r, column=col, value=f'FY{py}E')
        _style(c, fill_hex=MED_BLUE, bold=True, font_color=WHITE, h_align='center')
        c.border = THIN_BOX
    r += 1

    # -- Raw inputs (projection formulas backfilled after assumptions band) --
    rev_row  = r; _write_row(ws, r, 'Revenue', inc['revenue'], years); r += 1
    cogs_row = r; _write_row(ws, r, '  Cost of Revenue', inc['cogs'], years, indent=1); r += 1

    # Gross Profit = Revenue - Cost of Revenue  [FORMULA]
    gp_row = r
    _lbl(r, 'Gross Profit', bold=True)
    for i in range(n):
        _fw(r, i, f'={_cl(i)}{rev_row}-{_cl(i)}{cogs_row}', bold=True, fill=XLIGHT_BLUE)
    for j in range(n_proj):
        _pfw(r, j, f'={_pcl(j)}{rev_row}-{_pcl(j)}{cogs_row}', bold=True, fill=XLIGHT_BLUE)
    r += 1

    # Gross Margin % = Gross Profit / Revenue  [FORMULA]
    _lbl(r, '  Gross Margin %', ind=1)
    for i in range(n):
        _fw(r, i, f'=IF({_cl(i)}{rev_row}<>0,{_cl(i)}{gp_row}/{_cl(i)}{rev_row},"")',
            fmt=FMT_PCT)
    for j in range(n_proj):
        _pfw(r, j, f'=IF({_pcl(j)}{rev_row}<>0,{_pcl(j)}{gp_row}/{_pcl(j)}{rev_row},"")',
             fmt=FMT_PCT)
    r += 1

    _spacer(ws, r, total_cols); r += 1

    # -- Raw opex inputs (projection formulas backfilled after assumptions band) --
    rd_row  = r; _write_row(ws, r, '  R&D Expense', inc['rd_expense'], years, indent=1); r += 1
    sga_row = r; _write_row(ws, r, '  SG&A Expense', inc['sga_expense'], years, indent=1); r += 1

    # Other Operating Expenses / (Income): balancing plug so that EBIT
    # equals the exact reported Operating Income from the 10-K.
    # Plug = (Revenue - COGS) - R&D - SGA - Operating Income
    # Values in RAW DOLLARS — _write_row will scale to millions.
    other_opex_plug = {}
    for yr in years:
        gp_v   = (inc['revenue'].get(yr) or 0) - (inc['cogs'].get(yr) or 0)
        rd_v   = inc['rd_expense'].get(yr) or 0
        sga_v  = inc['sga_expense'].get(yr) or 0
        ebit_v = inc['operating_income'].get(yr)
        other_opex_plug[yr] = (gp_v - rd_v - sga_v - ebit_v) if ebit_v is not None else 0

    other_opex_row = r
    _write_row(ws, r, '  Other Operating Expenses / (Income)', other_opex_plug, years, indent=1)
    ws.cell(row=r, column=1).comment = Comment(
        "Plug: Gross Profit - R&D - SG&A - Operating Income.\n\n"
        "May include:\n"
        "- Depreciation & amortization\n"
        "- Restructuring charges\n"
        "- Impairment of assets\n"
        "- Acquisition-related costs\n"
        "- Litigation settlements\n"
        "- Other operating income/expense",
        "Financial Model")
    last_hist = _cl(n - 1)
    for j in range(n_proj):
        _pfw(r, j, f'={last_hist}{other_opex_row}')
    r += 1

    _spacer(ws, r, total_cols); r += 1

    # EBIT = Gross Profit - R&D - SGA - Other  [FORMULA]
    ebit_row = r
    _lbl(r, 'Operating Income (EBIT)', bold=True)
    for i in range(n):
        _fw(r, i,
            f'={_cl(i)}{gp_row}-{_cl(i)}{rd_row}-{_cl(i)}{sga_row}-{_cl(i)}{other_opex_row}',
            bold=True, fill=XLIGHT_BLUE)
    for j in range(n_proj):
        _pfw(r, j,
             f'={_pcl(j)}{gp_row}-{_pcl(j)}{rd_row}-{_pcl(j)}{sga_row}-{_pcl(j)}{other_opex_row}',
             bold=True, fill=XLIGHT_BLUE)
    r += 1

    # EBIT Margin % = EBIT / Revenue  [FORMULA]
    _lbl(r, '  EBIT Margin %', ind=1)
    for i in range(n):
        _fw(r, i,
            f'=IF({_cl(i)}{rev_row}<>0,{_cl(i)}{ebit_row}/{_cl(i)}{rev_row},"")',
            fmt=FMT_PCT)
    for j in range(n_proj):
        _pfw(r, j,
             f'=IF({_pcl(j)}{rev_row}<>0,{_pcl(j)}{ebit_row}/{_pcl(j)}{rev_row},"")',
             fmt=FMT_PCT)
    r += 1

    _spacer(ws, r, total_cols); r += 1

    # -- EBITDA add-backs --
    da_row = r
    _write_row(ws, r, '  Depreciation & Amortization', inc['da'], years, indent=1)
    for j in range(n_proj):
        _pfw(r, j, f'={last_hist}{da_row}')
    r += 1

    amort_row = r
    _write_row(ws, r, '  Amortization of Intangibles', inc['amortization'], years, indent=1)
    for j in range(n_proj):
        _pfw(r, j, f'={last_hist}{amort_row}')
    r += 1

    transform_row = r
    _write_row(ws, r, '  Transformation & Integration Costs', inc['transformation_costs'], years, indent=1)
    for j in range(n_proj):
        _pfw(r, j, f'={last_hist}{transform_row}')
    r += 1

    debt_ext_row = r
    _write_row(ws, r, '  Debt Extinguishment', inc['debt_extinguishment'], years, indent=1)
    for j in range(n_proj):
        _pfw(r, j, f'={last_hist}{debt_ext_row}')
    r += 1

    _spacer(ws, r, total_cols); r += 1

    # EBITDA = EBIT + D&A + Amort + Transform + |DebtExt|  [FORMULA]
    ebitda_row = r
    _lbl(r, 'EBITDA', bold=True)
    for i in range(n):
        _fw(r, i,
            f'={_cl(i)}{ebit_row}+{_cl(i)}{da_row}+{_cl(i)}{amort_row}'
            f'+{_cl(i)}{transform_row}+ABS({_cl(i)}{debt_ext_row})',
            bold=True, fill=XLIGHT_BLUE)
    for j in range(n_proj):
        _pfw(r, j,
             f'={_pcl(j)}{ebit_row}+{_pcl(j)}{da_row}+{_pcl(j)}{amort_row}'
             f'+{_pcl(j)}{transform_row}+ABS({_pcl(j)}{debt_ext_row})',
             bold=True, fill=XLIGHT_BLUE)
    r += 1

    # EBITDA Margin % = EBITDA / Revenue  [FORMULA]
    _lbl(r, '  EBITDA Margin %', ind=1)
    for i in range(n):
        _fw(r, i,
            f'=IF({_cl(i)}{rev_row}<>0,{_cl(i)}{ebitda_row}/{_cl(i)}{rev_row},"")',
            fmt=FMT_PCT)
    for j in range(n_proj):
        _pfw(r, j,
             f'=IF({_pcl(j)}{rev_row}<>0,{_pcl(j)}{ebitda_row}/{_pcl(j)}{rev_row},"")',
             fmt=FMT_PCT)
    r += 1

    # -- Stock-based Compensation (non-cash add-back) --
    sbc_row = r
    _write_row(ws, r, '  Stock-based Compensation', cf['sbc'], years, indent=1)
    for j in range(n_proj):
        _pfw(r, j, f'={last_hist}{sbc_row}')
    r += 1

    # Adjusted EBITDA = EBITDA + SBC  [FORMULA]
    adj_ebitda_row = r
    _lbl(r, 'Adjusted EBITDA', bold=True)
    for i in range(n):
        _fw(r, i, f'={_cl(i)}{ebitda_row}+{_cl(i)}{sbc_row}',
            bold=True, fill=XLIGHT_BLUE)
    for j in range(n_proj):
        _pfw(r, j, f'={_pcl(j)}{ebitda_row}+{_pcl(j)}{sbc_row}',
             bold=True, fill=XLIGHT_BLUE)
    r += 1

    _spacer(ws, r, total_cols); r += 1

    # -- Raw non-operating inputs + projections (hold flat) --
    int_exp_row = r; _write_row(ws, r, '  Interest Expense', inc['interest_expense'], years, indent=1)
    for j in range(n_proj):
        _pfw(r, j, f'={last_hist}{int_exp_row}')
    r += 1

    int_inc_row = r; _write_row(ws, r, '  Interest Income', inc['interest_income'], years, indent=1)
    for j in range(n_proj):
        _pfw(r, j, f'={last_hist}{int_inc_row}')
    r += 1

    other_inc_row = r; _write_row(ws, r, '  Other Income / (Expense)', inc['other_income'], years, indent=1)
    ws.cell(row=r, column=1).comment = Comment(
        "Plug: Pre-tax Income - EBIT - Interest Income"
        " + Interest Expense.\n\n"
        "May include:\n"
        "- Investment gains / losses\n"
        "- Equity method income\n"
        "- Foreign exchange gains / losses\n"
        "- Gains / losses on debt extinguishment\n"
        "- Other non-operating items",
        "Financial Model")
    for j in range(n_proj):
        _pfw(r, j, f'={last_hist}{other_inc_row}')
    r += 1

    # Pre-tax Income = EBIT - Interest Expense + Interest Income + Other Income  [FORMULA]
    pretax_row = r
    _lbl(r, 'Pre-tax Income', bold=True)
    for i in range(n):
        _fw(r, i,
            f'={_cl(i)}{ebit_row}-{_cl(i)}{int_exp_row}+{_cl(i)}{int_inc_row}+{_cl(i)}{other_inc_row}',
            bold=True, fill=XLIGHT_BLUE)
    for j in range(n_proj):
        _pfw(r, j,
             f'={_pcl(j)}{ebit_row}-{_pcl(j)}{int_exp_row}+{_pcl(j)}{int_inc_row}+{_pcl(j)}{other_inc_row}',
             bold=True, fill=XLIGHT_BLUE)
    r += 1

    # -- Tax (projection backfilled after assumptions band) --
    tax_row = r; _write_row(ws, r, '  Income Tax Expense', inc['tax_expense'], years, indent=1); r += 1

    # Effective Tax Rate = Tax / Pre-tax  [FORMULA]
    _lbl(r, '  Effective Tax Rate %', ind=1)
    for i in range(n):
        _fw(r, i,
            f'=IF({_cl(i)}{pretax_row}<>0,{_cl(i)}{tax_row}/{_cl(i)}{pretax_row},"")',
            fmt=FMT_PCT)
    for j in range(n_proj):
        _pfw(r, j,
             f'=IF({_pcl(j)}{pretax_row}<>0,{_pcl(j)}{tax_row}/{_pcl(j)}{pretax_row},"")',
             fmt=FMT_PCT)
    r += 1

    # Net Income = Pre-tax Income - Tax Expense  [FORMULA]
    ni_row = r
    _lbl(r, 'Net Income', bold=True)
    for i in range(n):
        _fw(r, i, f'={_cl(i)}{pretax_row}-{_cl(i)}{tax_row}', bold=True, fill=LIGHT_GREEN)
    for j in range(n_proj):
        _pfw(r, j, f'={_pcl(j)}{pretax_row}-{_pcl(j)}{tax_row}', bold=True, fill=LIGHT_GREEN)
    r += 1

    # Net Margin % = Net Income / Revenue  [FORMULA]
    _lbl(r, '  Net Margin %', ind=1)
    for i in range(n):
        _fw(r, i,
            f'=IF({_cl(i)}{rev_row}<>0,{_cl(i)}{ni_row}/{_cl(i)}{rev_row},"")',
            fmt=FMT_PCT)
    for j in range(n_proj):
        _pfw(r, j,
             f'=IF({_pcl(j)}{rev_row}<>0,{_pcl(j)}{ni_row}/{_pcl(j)}{rev_row},"")',
             fmt=FMT_PCT)
    r += 1

    _spacer(ws, r, total_cols); r += 1

    # EPS and shares are raw XBRL values (not in millions) — historical only
    _write_row(ws, r, '  EPS (Basic)',   inc['eps_basic'],   years, scale=1.0, fmt=FMT_DOLLAR2, indent=1); r += 1
    eps_diluted_row = r
    _write_row(ws, r, '  EPS (Diluted)', inc['eps_diluted'], years, scale=1.0, fmt=FMT_DOLLAR2, indent=1); r += 1
    _write_row(ws, r, '  Shares Outstanding - Basic (M)',   inc['shares_basic'],   years, scale=1e6, fmt=FMT_DOLLAR, indent=1); r += 1
    shares_diluted_row = r
    _write_row(ws, r, '  Shares Outstanding - Diluted (M)', inc['shares_diluted'], years, scale=1e6, fmt=FMT_DOLLAR, indent=1); r += 1
    r += 1

    # ── Growth Rates & Key Margins (assumptions band) ─────────────────
    # Written here (after all IS rows) so row numbers for Rev, GP, etc.
    # are already known.  Projection columns (F-J) filled later after
    # the scenario table at the bottom of the sheet.
    _lbl(r, 'GROWTH RATES & KEY MARGINS', bold=True, fill=LIGHT_BLUE)
    for ci in range(2, 2 + n + n_proj):
        ws.cell(row=r, column=ci).fill = _fill(LIGHT_BLUE)
    r += 1

    # Revenue Growth (YoY %) — N/A for oldest year
    rev_growth_asm_row = r; _lbl(r, '  Revenue Growth (YoY %)', ind=1)
    for i in range(n):
        if i == 0:
            c = ws.cell(row=r, column=2 + i, value='N/A')
            _style(c, fill_hex=MED_GRAY, h_align='center', italic=True)
        elif has_ann and years[i] == _ltm_info.get('ltm_year'):
            # Q cumulative: partial vs full year not comparable
            c = ws.cell(row=r, column=2 + i, value='N/A')
            _style(c, fill_hex=MED_GRAY, h_align='center', italic=True)
        elif has_ann and years[i] == ANN_YEAR:
            # Annualized growth vs latest actual annual year
            base_idx = years.index(_ltm_info['base_year'])
            _fw(r, i,
                f'=IF({_cl(base_idx)}{rev_row}<>0,{_cl(i)}{rev_row}/{_cl(base_idx)}{rev_row}-1,"")',
                fmt=FMT_PCT)
        else:
            _fw(r, i,
                f'=IF({_cl(i-1)}{rev_row}<>0,{_cl(i)}{rev_row}/{_cl(i-1)}{rev_row}-1,"")',
                fmt=FMT_PCT)
    r += 1

    # Gross Profit Margin (%)
    gp_margin_asm_row = r; _lbl(r, '  Gross Profit Margin (%)', ind=1)
    for i in range(n):
        _fw(r, i,
            f'=IF({_cl(i)}{rev_row}<>0,{_cl(i)}{gp_row}/{_cl(i)}{rev_row},"")',
            fmt=FMT_PCT)
    r += 1

    # R&D % of Sales
    rd_pct_asm_row = r; _lbl(r, '  R&D % of Sales', ind=1)
    for i in range(n):
        _fw(r, i,
            f'=IF({_cl(i)}{rev_row}<>0,{_cl(i)}{rd_row}/{_cl(i)}{rev_row},"")',
            fmt=FMT_PCT)
    r += 1

    # SG&A % of Sales
    sga_pct_asm_row = r; _lbl(r, '  SG&A % of Sales', ind=1)
    for i in range(n):
        _fw(r, i,
            f'=IF({_cl(i)}{rev_row}<>0,{_cl(i)}{sga_row}/{_cl(i)}{rev_row},"")',
            fmt=FMT_PCT)
    r += 1

    # Tax Rate (%)
    tax_rate_asm_row = r; _lbl(r, '  Tax Rate (%)', ind=1)
    for i in range(n):
        _fw(r, i,
            f'=IF({_cl(i)}{pretax_row}<>0,{_cl(i)}{tax_row}/{_cl(i)}{pretax_row},"")',
            fmt=FMT_PCT)
    r += 1

    # EBITDA Margin (%) — used by DCF via scenario dropdown
    ebitda_margin_asm_row = r; _lbl(r, '  EBITDA Margin (%)', ind=1)
    for i in range(n):
        _fw(r, i,
            f'=IF({_cl(i)}{rev_row}<>0,{_cl(i)}{ebitda_row}/{_cl(i)}{rev_row},"")',
            fmt=FMT_PCT)
    r += 1

    # D&A % of Revenue — used by DCF via scenario dropdown
    da_pct_asm_row = r; _lbl(r, '  D&A % of Revenue', ind=1)
    for i in range(n):
        _fw(r, i,
            f'=IF({_cl(i)}{rev_row}<>0,{_cl(i)}{da_row}/{_cl(i)}{rev_row},"")',
            fmt=FMT_PCT)
    r += 1

    # Capex % of Revenue — used by DCF via scenario dropdown
    # Historical: Python-computed (capex_row not yet defined at this point)
    capex_pct_asm_row = r; _lbl(r, '  Capex % of Revenue', ind=1)
    for i, yr in enumerate(years):
        cap_v = abs(cf['capex'].get(yr) or 0)
        rev_v = inc['revenue'].get(yr)
        if rev_v and rev_v != 0:
            val = cap_v / rev_v
        else:
            val = None
        if val is not None:
            c = ws.cell(row=r, column=2 + i, value=val)
            _style(c, h_align='right', number_format=FMT_PCT,
                   fill_hex=LIGHT_GRAY if i % 2 == 0 else WHITE)
        else:
            c = ws.cell(row=r, column=2 + i, value='N/A')
            _style(c, fill_hex=MED_GRAY, h_align='center', italic=True)
    r += 1

    _spacer(ws, r, total_cols); r += 1

    # ── Backfill IS projection formulas ─────────────────────────────
    # Now that assumptions band row numbers are known, fill in the
    # projection columns (F-J) for the IS line items above.
    for j in range(n_proj):
        prev = _cl(n - 1) if j == 0 else _pcl(j - 1)
        # Revenue = prior year × (1 + growth)
        _pfw(rev_row, j, f'={prev}{rev_row}*(1+{_pcl(j)}{rev_growth_asm_row})')
        # COGS = Revenue × (1 - GP margin)
        _pfw(cogs_row, j, f'={_pcl(j)}{rev_row}*(1-{_pcl(j)}{gp_margin_asm_row})')
        # R&D = Revenue × R&D %
        _pfw(rd_row, j, f'={_pcl(j)}{rev_row}*{_pcl(j)}{rd_pct_asm_row}')
        # SG&A = Revenue × SGA %
        _pfw(sga_row, j, f'={_pcl(j)}{rev_row}*{_pcl(j)}{sga_pct_asm_row}')
        # Tax = Pre-tax × Tax Rate
        _pfw(tax_row, j, f'={_pcl(j)}{pretax_row}*{_pcl(j)}{tax_rate_asm_row}')
        # SBC = prior year × (1 + revenue growth)
        _pfw(sbc_row, j, f'={prev}{sbc_row}*(1+{_pcl(j)}{rev_growth_asm_row})')

    # =========================================================================
    # BALANCE SHEET  (historical + projections)
    # =========================================================================
    _write_section_header(ws, r, 'BALANCE SHEET', cols=total_cols);  r += 1
    _write_col_headers(ws, r, list(range(2, 2 + n)), years, start_col=2)
    _fix_ann_headers(r)
    for j, py in enumerate(proj_years):
        col = proj_start_col + j
        c = ws.cell(row=r, column=col, value=f'FY{py}E')
        _style(c, fill_hex=MED_BLUE, bold=True, font_color=WHITE, h_align='center')
        c.border = THIN_BOX
    r += 1

    # Helper: growth-driven projection (prior year × (1 + rev growth))
    def _bs_grow(row_num):
        for j in range(n_proj):
            prev = _cl(n - 1) if j == 0 else _pcl(j - 1)
            _pfw(row_num, j, f'={prev}{row_num}*(1+{_pcl(j)}{rev_growth_asm_row})')

    # Helper: hold-flat projection (= last historical year)
    def _bs_flat(row_num):
        for j in range(n_proj):
            _pfw(row_num, j, f'={last_hist}{row_num}')

    _write_row(ws, r, 'ASSETS', {y: None for y in years}, years, bold=True); r += 1

    # Cash — backfilled later (will reference ending cash from CF)
    cash_row = r
    _write_row(ws, r, '  Cash & Cash Equivalents', bs['cash'], years, indent=1)
    _bs_flat(r)  # temporary: hold flat; will be overwritten after CF projections
    r += 1

    # ST Investments — grow with revenue
    st_inv_row = r
    _write_row(ws, r, '  Short-term Investments', bs['st_investments'], years, indent=1)
    _bs_grow(r); r += 1

    # Accounts Receivable — grow with revenue
    ar_row = r
    _write_row(ws, r, '  Accounts Receivable', bs['accounts_rec'], years, indent=1)
    _bs_grow(r); r += 1

    # Inventory — grows with COGS ratio
    inventory_row = r
    _write_row(ws, r, '  Inventory', bs['inventory'], years, indent=1)
    for j in range(n_proj):
        prev = _cl(n - 1) if j == 0 else _pcl(j - 1)
        _pfw(r, j,
             f'=IF({prev}{cogs_row}<>0,'
             f'{prev}{inventory_row}*{_pcl(j)}{cogs_row}/{prev}{cogs_row},'
             f'{prev}{inventory_row})')
    r += 1

    # Other Current Assets (plug) — grow with revenue
    other_ca_plug = {}
    for yr in years:
        tca = bs['total_current_a'].get(yr)
        if tca is not None:
            other_ca_plug[yr] = tca - sum(
                bs[k].get(yr) or 0
                for k in ('cash', 'st_investments', 'accounts_rec', 'inventory'))
        else:
            other_ca_plug[yr] = None
    other_ca_row = r
    _write_row(ws, r, '  Other Current Assets (plug)', other_ca_plug, years, indent=1)
    ws.cell(row=r, column=1).comment = Comment(
        "Plug: Total Current Assets - Cash - ST Investments"
        " - Accounts Receivable - Inventory.\n\n"
        "May include:\n"
        "- Prepaid expenses\n"
        "- Deferred tax assets (current)\n"
        "- Other receivables\n"
        "- Assets held for sale\n"
        "- Contract assets",
        "Financial Model")
    _bs_grow(r); r += 1

    # Total Current Assets [FORMULA for projections]
    total_ca_row = r
    _write_row(ws, r, 'Total Current Assets', bs['total_current_a'], years, bold=True, fill=XLIGHT_BLUE)
    for j in range(n_proj):
        cl = _pcl(j)
        _pfw(r, j,
             f'={cl}{cash_row}+{cl}{st_inv_row}+{cl}{ar_row}'
             f'+{cl}{inventory_row}+{cl}{other_ca_row}',
             bold=True, fill=XLIGHT_BLUE)
    r += 1

    _spacer(ws, r, total_cols); r += 1

    # PP&E — backfilled after PP&E schedule is written
    ppe_row = r
    _write_row(ws, r, '  PP&E, net', bs['ppe_net'], years, indent=1)
    _bs_flat(r)  # temporary: overwritten after PP&E schedule
    r += 1

    # Goodwill — grow with revenue
    gw_row = r
    _write_row(ws, r, '  Goodwill', bs['goodwill'], years, indent=1)
    _bs_grow(r); r += 1

    # Intangible Assets — grow with revenue
    intangibles_row = r
    _write_row(ws, r, '  Intangible Assets', bs['intangibles'], years, indent=1)
    _bs_grow(r); r += 1

    # Marketable Securities (non-current) — grow with revenue
    lt_inv_row = r
    _write_row(ws, r, '  Marketable Securities (non-current)', bs['lt_investments'], years, indent=1)
    _bs_grow(r); r += 1

    # Other Non-Current Assets (plug) — grow with revenue
    other_nca_plug = {}
    for yr in years:
        ta = bs['total_assets'].get(yr)
        tca = bs['total_current_a'].get(yr)
        if ta is not None and tca is not None:
            other_nca_plug[yr] = ta - tca - sum(
                bs[k].get(yr) or 0
                for k in ('ppe_net', 'goodwill', 'intangibles', 'lt_investments'))
        else:
            other_nca_plug[yr] = None
    other_nca_row = r
    _write_row(ws, r, '  Other Non-current Assets (plug)', other_nca_plug, years, indent=1)
    ws.cell(row=r, column=1).comment = Comment(
        "Plug: Total Assets - Total Current Assets - PP&E"
        " - Goodwill - Intangibles - Marketable Securities (non-current).\n\n"
        "May include:\n"
        "- Operating lease right-of-use assets\n"
        "- Deferred tax assets (non-current)\n"
        "- Non-current contract assets\n"
        "- Other long-term assets",
        "Financial Model")
    _bs_grow(r); r += 1

    # Total Assets [FORMULA for projections]
    total_assets_row = r
    _write_row(ws, r, 'Total Assets', bs['total_assets'], years, bold=True, fill=LIGHT_BLUE)
    for j in range(n_proj):
        cl = _pcl(j)
        _pfw(r, j,
             f'={cl}{total_ca_row}+{cl}{ppe_row}+{cl}{gw_row}'
             f'+{cl}{intangibles_row}+{cl}{lt_inv_row}+{cl}{other_nca_row}',
             bold=True, fill=LIGHT_BLUE)
    r += 1

    _spacer(ws, r, total_cols); r += 1

    # ── Liabilities ──────────────────────────────────────────────────
    _write_row(ws, r, 'LIABILITIES', {y: None for y in years}, years, bold=True); r += 1

    ap_row = r
    _write_row(ws, r, '  Accounts Payable', bs['accounts_pay'], years, indent=1)
    # Projection: prior year AP × (this year COGS / prior year COGS)
    for j in range(n_proj):
        prev = _cl(n - 1) if j == 0 else _pcl(j - 1)
        _pfw(r, j,
             f'=IF({prev}{cogs_row}<>0,'
             f'{prev}{ap_row}*{_pcl(j)}{cogs_row}/{prev}{cogs_row},'
             f'{prev}{ap_row})')
    r += 1

    accrued_row = r
    _write_row(ws, r, '  Accrued Liabilities', bs['accrued_liab'], years, indent=1)
    _bs_grow(r); r += 1

    st_debt_row = r
    _write_row(ws, r, '  Short-term Debt', bs['st_debt'], years, indent=1)
    _bs_grow(r); r += 1

    deferred_rev_row = r
    _write_row(ws, r, '  Deferred Revenue (current)', bs['deferred_rev_cur'], years, indent=1)
    _bs_grow(r); r += 1

    # Other Current Liabilities (plug) — grow with revenue
    other_cl_plug = {}
    for yr in years:
        tcl = bs['total_current_l'].get(yr)
        if tcl is not None:
            other_cl_plug[yr] = tcl - sum(
                bs[k].get(yr) or 0
                for k in ('accounts_pay', 'accrued_liab', 'st_debt', 'deferred_rev_cur'))
        else:
            other_cl_plug[yr] = None
    other_cl_row = r
    _write_row(ws, r, '  Other Current Liabilities (plug)', other_cl_plug, years, indent=1)
    ws.cell(row=r, column=1).comment = Comment(
        "Plug: Total Current Liabilities - Accounts Payable"
        " - Accrued Liabilities - ST Debt - Deferred Revenue.\n\n"
        "May include:\n"
        "- Operating lease liabilities (current)\n"
        "- Accrued income taxes\n"
        "- Dividends payable\n"
        "- Customer deposits\n"
        "- Other current liabilities",
        "Financial Model")
    _bs_grow(r); r += 1

    # Total Current Liabilities [FORMULA for projections]
    total_cl_row = r
    _write_row(ws, r, 'Total Current Liabilities', bs['total_current_l'], years, bold=True, fill=XLIGHT_BLUE)
    for j in range(n_proj):
        cl = _pcl(j)
        _pfw(r, j,
             f'={cl}{ap_row}+{cl}{accrued_row}+{cl}{st_debt_row}'
             f'+{cl}{deferred_rev_row}+{cl}{other_cl_row}',
             bold=True, fill=XLIGHT_BLUE)
    r += 1

    _spacer(ws, r, total_cols); r += 1

    lt_debt_row = r
    _write_row(ws, r, '  Long-term Debt', bs['lt_debt'], years, indent=1)
    _bs_grow(r); r += 1

    dtl_row = r
    _write_row(ws, r, '  Deferred Tax Liabilities', bs['deferred_tax_l'], years, indent=1)
    _bs_grow(r); r += 1

    # Other Non-Current Liabilities (plug) — grow with revenue
    other_ncl_plug = {}
    for yr in years:
        tl = bs['total_liabilities'].get(yr)
        tcl = bs['total_current_l'].get(yr)
        if tl is not None and tcl is not None:
            other_ncl_plug[yr] = tl - tcl - sum(
                bs[k].get(yr) or 0
                for k in ('lt_debt', 'deferred_tax_l'))
        else:
            other_ncl_plug[yr] = None
    other_ncl_row = r
    _write_row(ws, r, '  Other Non-current Liabilities (plug)', other_ncl_plug, years, indent=1)
    ws.cell(row=r, column=1).comment = Comment(
        "Plug: Total Liabilities - Total Current Liabilities"
        " - Long-term Debt - Deferred Tax Liabilities.\n\n"
        "May include:\n"
        "- Operating lease liabilities (non-current)\n"
        "- Finance lease liabilities\n"
        "- Pension & post-retirement obligations\n"
        "- Uncertain tax positions\n"
        "- Non-current deferred revenue\n"
        "- Other long-term liabilities",
        "Financial Model")
    _bs_grow(r); r += 1

    # Total Liabilities [FORMULA for projections]
    total_liabilities_row = r
    _write_row(ws, r, 'Total Liabilities', bs['total_liabilities'], years, bold=True, fill=LIGHT_BLUE)
    for j in range(n_proj):
        cl = _pcl(j)
        _pfw(r, j,
             f'={cl}{total_cl_row}+{cl}{lt_debt_row}+{cl}{dtl_row}+{cl}{other_ncl_row}',
             bold=True, fill=LIGHT_BLUE)
    r += 1

    _spacer(ws, r, total_cols); r += 1

    # ── Shareholders' Equity ─────────────────────────────────────────
    _write_row(ws, r, "SHAREHOLDERS' EQUITY", {y: None for y in years}, years, bold=True); r += 1

    cs_row = r
    _write_row(ws, r, '  Common Stock', bs['common_stock'], years, indent=1)
    _bs_flat(r); r += 1

    apic_row = r
    _write_row(ws, r, '  Additional Paid-in Capital', bs['apic'], years, indent=1)
    _bs_flat(r); r += 1

    re_row = r
    _write_row(ws, r, '  Retained Earnings', bs['retained_earnings'], years, indent=1)
    # Projection: prior year RE + Net Income
    for j in range(n_proj):
        prev = _cl(n - 1) if j == 0 else _pcl(j - 1)
        _pfw(r, j, f'={prev}{re_row}+{_pcl(j)}{ni_row}')
    r += 1

    ts_row = r
    _write_row(ws, r, '  Treasury Stock', bs['treasury_stock'], years, indent=1)
    _bs_flat(r); r += 1

    # Total Equity = Total Assets - Total Liabilities [FORMULA for projections]
    total_equity_row = r
    _write_row(ws, r, 'Total Equity', bs['total_equity'], years, bold=True, fill=XLIGHT_BLUE)
    for j in range(n_proj):
        cl = _pcl(j)
        _pfw(r, j, f'={cl}{total_assets_row}-{cl}{total_liabilities_row}',
             bold=True, fill=XLIGHT_BLUE)
    r += 1

    _spacer(ws, r, total_cols); r += 1

    # Total Liabilities + Equity  [FORMULA]
    _lbl(r, 'Total Liabilities + Equity', bold=True)
    for i in range(n):
        _fw(r, i,
            f'={_cl(i)}{total_liabilities_row}+{_cl(i)}{total_equity_row}',
            bold=True, fill=XLIGHT_BLUE)
    for j in range(n_proj):
        cl = _pcl(j)
        _pfw(r, j, f'={cl}{total_liabilities_row}+{cl}{total_equity_row}',
             bold=True, fill=XLIGHT_BLUE)
    r += 1

    # Balance Check: 0 = balanced, FALSE = imbalance
    _lbl(r, 'Balance Check  (0 = Balanced  |  FALSE = Imbalance)', bold=True)
    for i in range(n):
        a_cell = f'{_cl(i)}{total_assets_row}'
        diff   = f'ABS({a_cell}-({_cl(i)}{total_liabilities_row}+{_cl(i)}{total_equity_row}))'
        tol    = f'MAX({a_cell}*0.00001,0.5)'
        formula = f'=IF({diff}<{tol},0,FALSE)'
        c = ws.cell(row=r, column=2 + i, value=formula)
        _style(c, fill_hex=YELLOW, bold=True, h_align='center')
    # Projection balance check: always 0 by construction (TE = TA - TL)
    for j in range(n_proj):
        c = ws.cell(row=r, column=proj_start_col + j, value=0)
        _style(c, fill_hex=YELLOW, bold=True, h_align='center',
               number_format=FMT_DOLLAR)
    r += 2

    # =========================================================================
    # CASH FLOW STATEMENT  (historical + projections)
    # =========================================================================
    cf_cols = total_cols
    _write_section_header(ws, r, 'CASH FLOW STATEMENT', cols=cf_cols);  r += 1
    _write_col_headers(ws, r, list(range(2, 2 + n)), years, start_col=2)
    _fix_ann_headers(r)
    for j, py in enumerate(proj_years):
        col = proj_start_col + j
        c = ws.cell(row=r, column=col, value=f'FY{py}E')
        _style(c, fill_hex=MED_BLUE, bold=True, font_color=WHITE, h_align='center')
        c.border = THIN_BOX
    r += 1

    _write_row(ws, r, 'OPERATING ACTIVITIES', {y: None for y in years}, years, bold=True); r += 1

    # -- Net Income --
    cf_ni_row = r
    _write_row(ws, r, '  Net Income', cf['net_income'], years, indent=1)
    for j in range(n_proj):
        _pfw(r, j, f'={_pcl(j)}{ni_row}')
    r += 1

    # -- D&A --
    cf_da_row = r
    _write_row(ws, r, '  Depreciation & Amortization', cf['da'], years, indent=1)
    # Projection: placeholder — backfilled after PP&E schedule
    for j in range(n_proj):
        _pfw(r, j, f'={_pcl(j)}{da_row}')
    r += 1

    # -- SBC --
    cf_sbc_row = r
    _write_row(ws, r, '  Stock-based Compensation', cf['sbc'], years, indent=1)
    for j in range(n_proj):
        _pfw(r, j, f'={_pcl(j)}{sbc_row}')
    r += 1

    # -- Decreases / (Increases) in Working Capital Assets --
    # WC Assets = AR + Inventory + Other CA (excludes Cash, ST Investments)
    # Decrease in WC assets = source of cash (positive)
    cf_wc_a_row = r
    _lbl(r, '  Decreases / (Increases) in WC Assets', ind=1)
    for i in range(n):
        if i == 0:
            c = ws.cell(row=r, column=2 + i, value=0)
            _style(c, fill_hex=MED_GRAY, h_align='right', number_format=FMT_DOLLAR,
                   italic=True)
        elif has_ann and years[i] == ANN_YEAR:
            # Ann WC change: base_year → Q BS
            base_idx = years.index(_ltm_info['base_year'])
            prev = _cl(base_idx); curr = _cl(i)
            _fw(r, i,
                f'=({prev}{ar_row}+{prev}{inventory_row}+{prev}{other_ca_row})'
                f'-({curr}{ar_row}+{curr}{inventory_row}+{curr}{other_ca_row})')
        else:
            prev = _cl(i - 1); curr = _cl(i)
            _fw(r, i,
                f'=({prev}{ar_row}+{prev}{inventory_row}+{prev}{other_ca_row})'
                f'-({curr}{ar_row}+{curr}{inventory_row}+{curr}{other_ca_row})')
    for j in range(n_proj):
        prev = _cl(n - 1) if j == 0 else _pcl(j - 1)
        curr = _pcl(j)
        _pfw(r, j,
             f'=({prev}{ar_row}+{prev}{inventory_row}+{prev}{other_ca_row})'
             f'-({curr}{ar_row}+{curr}{inventory_row}+{curr}{other_ca_row})')
    r += 1

    # -- Increases / (Decreases) in Working Capital Liabilities --
    # WC Liabilities = AP + Accrued + Deferred Rev + Other CL (excludes ST Debt)
    # Increase in WC liabilities = source of cash (positive)
    cf_wc_l_row = r
    _lbl(r, '  Increases / (Decreases) in WC Liabilities', ind=1)
    for i in range(n):
        if i == 0:
            c = ws.cell(row=r, column=2 + i, value=0)
            _style(c, fill_hex=MED_GRAY, h_align='right', number_format=FMT_DOLLAR,
                   italic=True)
        elif has_ann and years[i] == ANN_YEAR:
            base_idx = years.index(_ltm_info['base_year'])
            prev = _cl(base_idx); curr = _cl(i)
            _fw(r, i,
                f'=({curr}{ap_row}+{curr}{accrued_row}+{curr}{deferred_rev_row}+{curr}{other_cl_row})'
                f'-({prev}{ap_row}+{prev}{accrued_row}+{prev}{deferred_rev_row}+{prev}{other_cl_row})')
        else:
            prev = _cl(i - 1); curr = _cl(i)
            _fw(r, i,
                f'=({curr}{ap_row}+{curr}{accrued_row}+{curr}{deferred_rev_row}+{curr}{other_cl_row})'
                f'-({prev}{ap_row}+{prev}{accrued_row}+{prev}{deferred_rev_row}+{prev}{other_cl_row})')
    for j in range(n_proj):
        prev = _cl(n - 1) if j == 0 else _pcl(j - 1)
        curr = _pcl(j)
        _pfw(r, j,
             f'=({curr}{ap_row}+{curr}{accrued_row}+{curr}{deferred_rev_row}+{curr}{other_cl_row})'
             f'-({prev}{ap_row}+{prev}{accrued_row}+{prev}{deferred_rev_row}+{prev}{other_cl_row})')
    r += 1

    # -- Other Operating Activities (plug) --
    # Historical: backs into actual Cash from Ops; Projection: 0
    cf_other_op_row = r
    _lbl(r, '  Other Operating Activities (plug)', ind=1)
    # Pre-compute op_cf_row position for formula (comes 1 row after this)
    op_cf_row = r + 1
    for i in range(n):
        cl = _cl(i)
        _fw(r, i,
            f'={cl}{op_cf_row}-{cl}{cf_ni_row}-{cl}{cf_da_row}'
            f'-{cl}{cf_sbc_row}-{cl}{cf_wc_a_row}-{cl}{cf_wc_l_row}')
    ws.cell(row=r, column=1).comment = Comment(
        "Plug: Cash from Operations - Net Income - D&A - SBC"
        " - WC Asset Changes - WC Liability Changes.\n\n"
        "May include:\n"
        "- Deferred income taxes\n"
        "- Impairment / write-down charges\n"
        "- Gains/losses on investments\n"
        "- Amortization of debt issuance costs\n"
        "- Other non-cash adjustments",
        "Financial Model")
    for j in range(n_proj):
        _pfw(r, j, f'=0')
    r += 1

    # Cash from Operations
    # Historical: actual data; Projection: SUM of operating items
    _write_row(ws, r, 'Cash from Operations', cf['operating_cf'], years,
               bold=True, fill=LIGHT_GREEN)
    for j in range(n_proj):
        cl = _pcl(j)
        _pfw(r, j,
             f'={cl}{cf_ni_row}+{cl}{cf_da_row}+{cl}{cf_sbc_row}'
             f'+{cl}{cf_wc_a_row}+{cl}{cf_wc_l_row}+{cl}{cf_other_op_row}',
             bold=True, fill=LIGHT_GREEN)
    r += 1

    _spacer(ws, r, cf_cols); r += 1

    # -- INVESTING ACTIVITIES --
    _write_row(ws, r, 'INVESTING ACTIVITIES', {y: None for y in years}, years, bold=True); r += 1
    capex_row = r
    _write_row(ws, r, '  Capital Expenditures (Capex)', cf['capex'], years,
               negate=True, indent=1)
    # Projection: placeholder — backfilled after PP&E schedule
    for j in range(n_proj):
        _pfw(r, j, f'=0')
    r += 1

    _write_row(ws, r, '  Acquisitions', cf['acquisitions'], years,
               negate=True, indent=1)
    for j in range(n_proj):
        _pfw(r, j, f'=0')
    r += 1

    # Other Investing Activities: plug for historical, 0 for projections
    other_inv_plug = {}
    for yr in years:
        inv = cf['investing_cf'].get(yr)
        if inv is not None:
            other_inv_plug[yr] = inv + sum(
                cf[k].get(yr) or 0
                for k in ('capex', 'acquisitions'))
        else:
            other_inv_plug[yr] = None
    _write_row(ws, r, '  Other Investing Activities (plug)', other_inv_plug,
               years, indent=1)
    ws.cell(row=r, column=1).comment = Comment(
        "Plug: Cash from Investing + Capex + Acquisitions.\n\n"
        "May include:\n"
        "- Purchases of marketable securities\n"
        "- Maturities / sales of marketable securities\n"
        "- Purchases of non-marketable securities\n"
        "- Other investing activities",
        "Financial Model")
    for j in range(n_proj):
        _pfw(r, j, f'=0')
    r += 1

    inv_cf_row = r
    _write_row(ws, r, 'Cash from Investing', cf['investing_cf'], years,
               bold=True, fill=LIGHT_RED)
    # Projection: = Capex (already negative) + Acq + Other
    for j in range(n_proj):
        _pfw(r, j, f'={_pcl(j)}{capex_row}', bold=True, fill=LIGHT_RED)
    r += 1

    _spacer(ws, r, cf_cols); r += 1

    # -- FINANCING ACTIVITIES --
    _write_row(ws, r, 'FINANCING ACTIVITIES', {y: None for y in years}, years, bold=True); r += 1

    # Dividends Paid (historical negative via negate=True in _write_row)
    # Projection: -MIN(|prior year dividends|, |this year net income|)
    dividends_row = r
    _write_row(ws, r, '  Dividends Paid', cf['dividends'], years, negate=True, indent=1)
    for j in range(n_proj):
        prev = _cl(n - 1) if j == 0 else _pcl(j - 1)
        cl = _pcl(j)
        _pfw(r, j, f'=-MIN(ABS({prev}{dividends_row}),ABS({cl}{ni_row}))')
    r += 1

    # Share Repurchases (historical negative via negate=True)
    # Projection: placeholder — backfilled to reference RE schedule
    repurchases_row = r
    _write_row(ws, r, '  Share Repurchases', cf['repurchases'], years, negate=True, indent=1)
    for j in range(n_proj):
        _pfw(r, j, f'=0')  # placeholder, backfilled after RE schedule
    r += 1

    # Debt Issuance (positive = cash inflow)
    # Projection: prior year × (1 + revenue growth)
    debt_iss_row = r
    _write_row(ws, r, '  Debt Issuance', cf['debt_issuance'], years, indent=1)
    for j in range(n_proj):
        prev = _cl(n - 1) if j == 0 else _pcl(j - 1)
        cl = _pcl(j)
        _pfw(r, j, f'={prev}{debt_iss_row}*(1+{cl}{rev_growth_asm_row})')
    r += 1

    # Debt Repayment (historical negative via negate=True)
    # Projection: -MIN(|prior year repay| × (1 + growth), ST debt + LT debt)
    debt_rep_row = r
    _write_row(ws, r, '  Debt Repayment', cf['debt_repay'], years, negate=True, indent=1)
    for j in range(n_proj):
        prev = _cl(n - 1) if j == 0 else _pcl(j - 1)
        cl = _pcl(j)
        _pfw(r, j,
             f'=-MIN(ABS({prev}{debt_rep_row})*(1+{cl}{rev_growth_asm_row}),'
             f'{cl}{st_debt_row}+{cl}{lt_debt_row})')
    r += 1

    # Other Financing Activities (plug): historical only, projection = 0
    other_fin_plug = {}
    for yr in years:
        fin = cf['financing_cf'].get(yr)
        if fin is not None:
            other_fin_plug[yr] = (fin
                + (cf['dividends'].get(yr) or 0)
                + (cf['repurchases'].get(yr) or 0)
                - (cf['debt_issuance'].get(yr) or 0)
                + (cf['debt_repay'].get(yr) or 0))
        else:
            other_fin_plug[yr] = None
    other_fin_row = r
    _write_row(ws, r, '  Other Financing Activities (plug)', other_fin_plug,
               years, indent=1)
    ws.cell(row=r, column=1).comment = Comment(
        "Plug: Cash from Financing + Dividends + Repurchases"
        " - Debt Issuance + Debt Repayment.\n\n"
        "May include:\n"
        "- Stock option / RSU proceeds\n"
        "- Finance lease principal payments\n"
        "- Contingent consideration payments\n"
        "- Other financing activities",
        "Financial Model")
    for j in range(n_proj):
        _pfw(r, j, f'=0')
    r += 1

    # Cash from Financing = SUM of all financing items
    # All items already have correct sign (dividends/repurchases/debt_rep negative,
    # debt_issuance positive, other = 0)
    fin_cf_row = r
    _write_row(ws, r, 'Cash from Financing', cf['financing_cf'], years,
               bold=True, fill=LIGHT_RED)
    for j in range(n_proj):
        cl = _pcl(j)
        _pfw(r, j,
             f'={cl}{dividends_row}+{cl}{repurchases_row}'
             f'+{cl}{debt_iss_row}+{cl}{debt_rep_row}+{cl}{other_fin_row}',
             bold=True, fill=LIGHT_RED)
    r += 1

    _spacer(ws, r, cf_cols); r += 1

    fx_cf_row = r
    _write_row(ws, r, 'Effect of Exchange Rate Changes on Cash', cf['fx_effect'],
               years)
    for j in range(n_proj):
        _pfw(r, j, f'=0')
    r += 1

    _spacer(ws, r, cf_cols); r += 1

    fcf_row = r
    _lbl(r, 'Free Cash Flow  (Op + Inv + Fin + FX)', bold=True)
    for i in range(n):
        _fw(r, i,
            f'={_cl(i)}{op_cf_row}+{_cl(i)}{inv_cf_row}+{_cl(i)}{fin_cf_row}+{_cl(i)}{fx_cf_row}',
            bold=True, fill=LIGHT_GREEN)
    for j in range(n_proj):
        _pfw(r, j,
             f'={_pcl(j)}{op_cf_row}+{_pcl(j)}{inv_cf_row}+{_pcl(j)}{fin_cf_row}+{_pcl(j)}{fx_cf_row}',
             bold=True, fill=LIGHT_GREEN)
    r += 1

    # FCF Margin %
    _lbl(r, '  FCF Margin %', ind=1)
    for i in range(n):
        _fw(r, i,
            f'=IF({_cl(i)}{rev_row}<>0,{_cl(i)}{fcf_row}/{_cl(i)}{rev_row},"")',
            fmt=FMT_PCT)
    for j in range(n_proj):
        _pfw(r, j,
             f'=IF({_pcl(j)}{rev_row}<>0,{_pcl(j)}{fcf_row}/{_pcl(j)}{rev_row},"")',
             fmt=FMT_PCT)
    r += 2

    # =========================================================================
    # CASH FLOW RECONCILIATION
    # Verifies: Beginning Cash + Net Change in Cash = Ending Cash on Balance Sheet
    # The check cell returns 0 if the CF statement ties to the BS; FALSE otherwise.
    # (First year has no prior-year data, so only years 2 & 3 are checkable.)
    # =========================================================================
    _write_section_header(
        ws, r,
        'CASH FLOW RECONCILIATION  (Beginning Cash + Net Change = Ending Cash on BS)',
        cols=cf_cols)
    r += 1
    # Column headers: historical + projection
    _write_col_headers(ws, r, list(range(2, 2 + n)), years, start_col=2)
    _fix_ann_headers(r)
    for j, py in enumerate(proj_years):
        col = proj_start_col + j
        c = ws.cell(row=r, column=col, value=f'FY{py}E')
        _style(c, fill_hex=MED_BLUE, bold=True, font_color=WHITE, h_align='center')
        c.border = THIN_BOX
    r += 1

    # Beginning Cash = prior year's ending cash
    beg_cash_row = r
    _lbl(r, 'Beginning Cash  (prior year Ending Cash)')
    for i in range(n):
        if i == 0:
            c = ws.cell(row=r, column=2 + i, value='N/A (no prior year)')
            _style(c, fill_hex=MED_GRAY, h_align='center', italic=True)
        elif has_ann and years[i] == ANN_YEAR:
            # Ann beginning cash = base_year cash
            base_idx = years.index(_ltm_info['base_year'])
            _fw(r, i, f'={_cl(base_idx)}{cash_row}')
        else:
            prev = get_column_letter(2 + i - 1)
            _fw(r, i, f'={prev}{cash_row}')
    # Projection: beginning cash = prior year's ending cash (from this recon)
    # end_cf_row is defined below — use forward-known offset
    # We know the row layout: beg_cash, +4 CF rows, net_chg, end_cash
    # end_cf_row = beg_cash_row + 6
    # For j=0, beginning cash = last historical year BS cash
    # For j>0, beginning cash = prior projection ending cash
    for j in range(n_proj):
        if j == 0:
            _pfw(r, j, f'={last_hist}{cash_row}')
        else:
            # Forward reference to end_cf_row — will be at beg_cash_row + 6
            _pfw(r, j, f'={_pcl(j-1)}{beg_cash_row + 6}')
    r += 1

    # CF section totals (referenced from CF statement above)
    recon_cf_rows = []
    for label, src in [
        ('+ Cash from Operations',              op_cf_row),
        ('+ Cash from Investing',               inv_cf_row),
        ('+ Cash from Financing',               fin_cf_row),
        ('+ Effect of FX on Cash',              fx_cf_row),
    ]:
        recon_cf_rows.append(r)
        _lbl(r, label, ind=1)
        for i in range(n):
            _fw(r, i, f'={_cl(i)}{src}')
        for j in range(n_proj):
            _pfw(r, j, f'={_pcl(j)}{src}')
        r += 1

    # Net Change in Cash = Op + Inv + Fin + FX  [FORMULA]
    net_chg_row = r
    _lbl(r, 'Net Change in Cash', bold=True)
    for i in range(n):
        _fw(r, i,
            f'={_cl(i)}{op_cf_row}+{_cl(i)}{inv_cf_row}+{_cl(i)}{fin_cf_row}+{_cl(i)}{fx_cf_row}',
            bold=True, fill=XLIGHT_BLUE)
    for j in range(n_proj):
        cl = _pcl(j)
        _pfw(r, j,
             f'={cl}{op_cf_row}+{cl}{inv_cf_row}+{cl}{fin_cf_row}+{cl}{fx_cf_row}',
             bold=True, fill=XLIGHT_BLUE)
    r += 1

    # Ending Cash = Beginning + Net Change  [FORMULA]
    end_cf_row = r
    assert end_cf_row == beg_cash_row + 6, \
        f"end_cf_row layout assumption broken: {end_cf_row} != {beg_cash_row + 6}"
    _lbl(r, 'Ending Cash  (CF-derived)', bold=True)
    for i in range(n):
        if i == 0:
            c = ws.cell(row=r, column=2 + i, value='N/A')
            _style(c, fill_hex=MED_GRAY, h_align='center', italic=True)
        else:
            _fw(r, i, f'={_cl(i)}{beg_cash_row}+{_cl(i)}{net_chg_row}',
                bold=True, fill=LIGHT_BLUE)
    for j in range(n_proj):
        _pfw(r, j, f'={_pcl(j)}{beg_cash_row}+{_pcl(j)}{net_chg_row}',
             bold=True, fill=LIGHT_BLUE)
    r += 1

    # Ending Cash per Balance Sheet  [FORMULA, references BS cash row]
    _lbl(r, 'Ending Cash  (Balance Sheet)', bold=True)
    for i in range(n):
        _fw(r, i, f'={_cl(i)}{cash_row}', bold=True, fill=LIGHT_BLUE)
    for j in range(n_proj):
        _pfw(r, j, f'={_pcl(j)}{cash_row}', bold=True, fill=LIGHT_BLUE)
    r += 1

    # Reconciliation Check: 0 = CF ties to BS, FALSE = discrepancy  [FORMULA]
    # Historical: tolerance-based check for XBRL rounding
    # Projection: always 0 by construction (BS cash = ending cash from recon)
    _lbl(r, 'Reconciliation Check  (0 = OK  |  FALSE = Mismatch)', bold=True)
    for i in range(n):
        if i == 0:
            c = ws.cell(row=r, column=2 + i, value='N/A')
            _style(c, fill_hex=MED_GRAY, bold=True, h_align='center')
        else:
            diff = f'ABS({_cl(i)}{end_cf_row}-{_cl(i)}{cash_row})'
            tol  = f'MAX(ABS({_cl(i)}{cash_row})*0.00001,0.5)'
            formula = f'=IF({diff}<{tol},0,FALSE)'
            c = ws.cell(row=r, column=2 + i, value=formula)
            _style(c, fill_hex=YELLOW, bold=True, h_align='center')
    for j in range(n_proj):
        c = ws.cell(row=r, column=proj_start_col + j, value=0)
        _style(c, fill_hex=YELLOW, bold=True, h_align='center',
               number_format=FMT_DOLLAR)
    r += 2

    # =========================================================================
    # PP&E SCHEDULE
    # =========================================================================
    _write_section_header(ws, r, 'PP&E SCHEDULE', cols=total_cols); r += 1
    # Column headers: historical + projection
    _write_col_headers(ws, r, list(range(2, 2 + n)), years, start_col=2)
    _fix_ann_headers(r)
    for j, py in enumerate(proj_years):
        col = proj_start_col + j
        c = ws.cell(row=r, column=col, value=f'FY{py}E')
        _style(c, fill_hex=MED_BLUE, bold=True, font_color=WHITE, h_align='center')
        c.border = THIN_BOX
    # Step-up column header
    step_col = total_cols + 1
    step_col_letter = get_column_letter(step_col)
    c = ws.cell(row=r, column=step_col, value='Step-up')
    _style(c, fill_hex=DARK_BLUE, bold=True, font_color=WHITE, h_align='center')
    c.border = THIN_BOX
    ws.column_dimensions[step_col_letter].width = 14
    r += 1

    # Pre-compute row positions so all formulas can cross-reference
    ppe_beg_sch_row    = r
    ppe_capex_sch_row  = r + 1
    ppe_depr_sch_row   = r + 2
    ppe_end_sch_row    = r + 3
    # r + 4 = spacer
    da_capex_pct_row   = r + 5
    capex_pct_rev_row  = r + 6

    step_up_ref   = f'${step_col_letter}${da_capex_pct_row}'
    capex_pct_ref = f'${step_col_letter}${capex_pct_rev_row}'

    # --- Beginning PP&E ---
    _lbl(r, 'Beginning PP&E')
    for i in range(n):
        if i == 0:
            c = ws.cell(row=r, column=2 + i, value='N/A')
            _style(c, fill_hex=MED_GRAY, h_align='center', italic=True)
        elif has_ann and years[i] == ANN_YEAR:
            # Ann beginning = base_year PP&E
            base_idx = years.index(_ltm_info['base_year'])
            _fw(r, i, f'={_cl(base_idx)}{ppe_row}')
        else:
            _fw(r, i, f'={_cl(i-1)}{ppe_row}')
    for j in range(n_proj):
        if j == 0:
            _pfw(r, j, f'={last_hist}{ppe_row}')
        else:
            _pfw(r, j, f'={_pcl(j-1)}{ppe_end_sch_row}')
    r += 1

    # --- + Capital Expenditures ---
    _lbl(r, '+ Capital Expenditures')
    for i in range(n):
        _fw(r, i, f'=ABS({_cl(i)}{capex_row})')
    _capex_ref_yr = ANN_YEAR if has_ann else latest_yr
    last_hist_capex_val = abs(_val(cf['capex'], _capex_ref_yr) or 0)
    for j in range(n_proj):
        if j < 2:
            # Editable yellow cells — default to last historical capex
            c = ws.cell(row=r, column=proj_start_col + j,
                        value=round(last_hist_capex_val, 1))
            _style(c, fill_hex=YELLOW, h_align='right', number_format=FMT_DOLLAR)
            c.border = THIN_BOX
        else:
            # Years 3-5: Revenue × Capex % of Revenue
            _pfw(r, j, f'={_pcl(j)}{rev_row}*{capex_pct_ref}')
    r += 1

    # --- - Depreciation ---
    _lbl(r, '- Depreciation')
    for i in range(n):
        _fw(r, i, f'={_cl(i)}{da_row}')
    for j in range(n_proj):
        _pfw(r, j, f'={_pcl(j)}{ppe_capex_sch_row}*{_pcl(j)}{da_capex_pct_row}')
    r += 1

    # --- Ending PP&E = Beg + Capex - Depr ---
    _lbl(r, 'Ending PP&E', bold=True)
    for i in range(n):
        # Historical: anchor to BS PP&E value (absorbs disposals, impairments)
        _fw(r, i, f'={_cl(i)}{ppe_row}', bold=True, fill=XLIGHT_BLUE)
    for j in range(n_proj):
        cl = _pcl(j)
        _pfw(r, j,
             f'={cl}{ppe_beg_sch_row}+{cl}{ppe_capex_sch_row}-{cl}{ppe_depr_sch_row}',
             bold=True, fill=XLIGHT_BLUE)
    r += 1

    _spacer(ws, r, total_cols); r += 1

    # --- D&A / Capex % ---
    _lbl(r, 'D&A / Capex %')
    for i in range(n):
        _fw(r, i,
            f'=IF({_cl(i)}{ppe_capex_sch_row}<>0,'
            f'{_cl(i)}{ppe_depr_sch_row}/{_cl(i)}{ppe_capex_sch_row},"")',
            fmt=FMT_PCT)
    # Yr 1: average of last 3 historical; Yr 2+: prior + step-up
    avg_start = max(0, n - 3)
    avg_cells = ','.join(f'{_cl(i)}{da_capex_pct_row}' for i in range(avg_start, n))
    for j in range(n_proj):
        if j == 0:
            _pfw(r, j, f'=AVERAGE({avg_cells})', fmt=FMT_PCT)
        else:
            _pfw(r, j, f'={_pcl(j-1)}{da_capex_pct_row}+{step_up_ref}', fmt=FMT_PCT)
    # Step-up editable cell (default 2.0%)
    c = ws.cell(row=r, column=step_col, value=0.02)
    _style(c, fill_hex=YELLOW, h_align='right', number_format=FMT_PCT)
    c.border = THIN_BOX
    r += 1

    # --- Capex % of Revenue (for projection years 3-5) ---
    _lbl(r, 'Capex % of Revenue (Yr 3-5)')
    for i in range(n):
        _fw(r, i,
            f'=IF({_cl(i)}{rev_row}<>0,'
            f'ABS({_cl(i)}{capex_row})/{_cl(i)}{rev_row},"")',
            fmt=FMT_PCT)
    # Editable yellow cell in step-up column — default to historical average
    capex_rev_vals = []
    _excl_yrs = set()
    if has_ann:
        _excl_yrs.add(_ltm_info.get('ltm_year'))
        _excl_yrs.add(ANN_YEAR)
    for yr in years:
        if yr in _excl_yrs:
            continue
        cap_v = abs((_val(cf['capex'], yr) or 0))
        rev_v = _val(inc['revenue'], yr)
        if cap_v > 0 and rev_v and rev_v > 0:
            capex_rev_vals.append(cap_v / rev_v)
    avg_capex_pct_rev = _safe_avg(capex_rev_vals) if capex_rev_vals else 0.05
    c = ws.cell(row=r, column=step_col, value=round(avg_capex_pct_rev, 4))
    _style(c, fill_hex=YELLOW, h_align='right', number_format=FMT_PCT)
    c.border = THIN_BOX
    r += 2

    # =========================================================================
    # RETAINED EARNINGS SCHEDULE
    # =========================================================================
    _write_section_header(ws, r, 'RETAINED EARNINGS SCHEDULE', cols=total_cols); r += 1
    # Column headers: historical + projection
    _write_col_headers(ws, r, list(range(2, 2 + n)), years, start_col=2)
    _fix_ann_headers(r)
    for j, py in enumerate(proj_years):
        col = proj_start_col + j
        c = ws.cell(row=r, column=col, value=f'FY{py}E')
        _style(c, fill_hex=MED_BLUE, bold=True, font_color=WHITE, h_align='center')
        c.border = THIN_BOX
    r += 1

    # Pre-compute row positions
    re_beg_sch_row  = r
    re_ni_sch_row   = r + 1
    re_div_sch_row  = r + 2
    re_rep_sch_row  = r + 3
    re_end_sch_row  = r + 4

    # --- Beginning Retained Earnings ---
    _lbl(r, 'Beginning Retained Earnings')
    for i in range(n):
        if i == 0:
            c = ws.cell(row=r, column=2 + i, value='N/A')
            _style(c, fill_hex=MED_GRAY, h_align='center', italic=True)
        elif has_ann and years[i] == ANN_YEAR:
            base_idx = years.index(_ltm_info['base_year'])
            _fw(r, i, f'={_cl(base_idx)}{re_row}')
        else:
            _fw(r, i, f'={_cl(i-1)}{re_row}')
    for j in range(n_proj):
        if j == 0:
            _pfw(r, j, f'={last_hist}{re_row}')
        else:
            _pfw(r, j, f'={_pcl(j-1)}{re_end_sch_row}')
    r += 1

    # --- + Net Income ---
    _lbl(r, '+ Net Income')
    for i in range(n):
        _fw(r, i, f'={_cl(i)}{ni_row}')
    for j in range(n_proj):
        _pfw(r, j, f'={_pcl(j)}{ni_row}')
    r += 1

    # --- - Dividends ---
    # CF dividends_row: stored positive (absolute), displayed negative via negate
    # In projections, dividends_row cells are already negative
    _lbl(r, '- Dividends')
    for i in range(n):
        _fw(r, i, f'=ABS({_cl(i)}{dividends_row})')
    for j in range(n_proj):
        _pfw(r, j, f'=ABS({_pcl(j)}{dividends_row})')
    r += 1

    # --- - Share Repurchases ---
    # Historical: actual repurchases from CF; Projection: previous year's amount
    _lbl(r, '- Share Repurchases')
    for i in range(n):
        _fw(r, i, f'=ABS({_cl(i)}{repurchases_row})')
    for j in range(n_proj):
        prev = _cl(n - 1) if j == 0 else _pcl(j - 1)
        _pfw(r, j, f'=ABS({prev}{repurchases_row})')
    r += 1

    # --- Ending Retained Earnings = Beg + NI - Div - Repurchases ---
    _lbl(r, 'Ending Retained Earnings', bold=True)
    for i in range(n):
        # Historical: anchor to BS RE value (absorbs other comprehensive income etc.)
        _fw(r, i, f'={_cl(i)}{re_row}', bold=True, fill=XLIGHT_BLUE)
    for j in range(n_proj):
        cl = _pcl(j)
        _pfw(r, j,
             f'={cl}{re_beg_sch_row}+{cl}{re_ni_sch_row}'
             f'-{cl}{re_div_sch_row}-{cl}{re_rep_sch_row}',
             bold=True, fill=XLIGHT_BLUE)
    r += 2

    # ── Backfill: BS Retained Earnings → RE schedule ending balance ──
    for j in range(n_proj):
        _pfw(re_row, j, f'={_pcl(j)}{re_end_sch_row}')

    # ── Backfill: CF Share Repurchases → negated RE schedule repurchases ─
    for j in range(n_proj):
        _pfw(repurchases_row, j, f'=-{_pcl(j)}{re_rep_sch_row}')

    # ── Backfill: BS PP&E projections → PP&E schedule ending balance ──
    for j in range(n_proj):
        _pfw(ppe_row, j, f'={_pcl(j)}{ppe_end_sch_row}')

    # ── Backfill: IS D&A projections → PP&E schedule depreciation ─────
    # Override the "hold flat" D&A projections with PP&E-schedule-driven values
    for j in range(n_proj):
        _pfw(da_row, j, f'={_pcl(j)}{ppe_depr_sch_row}')

    # ── Backfill: CF D&A projections → PP&E schedule depreciation ─────
    for j in range(n_proj):
        _pfw(cf_da_row, j, f'={_pcl(j)}{ppe_depr_sch_row}')

    # ── Backfill: CF Capex projections → PP&E schedule capex ──────────
    for j in range(n_proj):
        _pfw(capex_row, j, f'=-{_pcl(j)}{ppe_capex_sch_row}')

    # ── Backfill: BS Cash projections → CF Reconciliation ending cash ─
    for j in range(n_proj):
        _pfw(cash_row, j, f'={_pcl(j)}{end_cf_row}')

    ws.freeze_panes = 'A3'

    # =========================================================================
    # 5-YEAR REVENUE CAGR REFERENCE
    # =========================================================================
    _ext_rev = fd.get('revenue_extended', {})
    _ext_yrs = sorted([y for y, v in _ext_rev.items() if v is not None and v > 0])
    _latest_hist_cl = get_column_letter(2 + n - 1)

    # Use only actual annual years (exclude Q cum / annualized) for CAGR
    _annual_yrs = [yr for yr in years if isinstance(yr, int) and
                   (not has_ann or yr != _ltm_info.get('ltm_year'))]

    if len(_ext_yrs) >= 2:
        _cagr_start_year = _ext_yrs[0]
        _cagr_start_rev = _ext_rev[_cagr_start_year] / 1e6
        _cagr_periods = latest_yr - _cagr_start_year
    else:
        _cagr_start_year = _annual_yrs[0] if _annual_yrs else years[0]
        _cagr_start_rev = (inc['revenue'].get(_cagr_start_year) or 0) / 1e6
        _cagr_periods = latest_yr - _cagr_start_year if len(_annual_yrs) > 1 else 1

    if _cagr_periods < 1:
        _cagr_periods = 1

    _cagr_end_label = (_ltm_info['ann_label'] if has_ann
                       else f'FY{latest_yr}')

    r += 2
    _write_section_header(ws, r, '5-YEAR REVENUE CAGR', cols=total_cols); r += 1

    _lbl(r, f'Starting Revenue ($M, FY{_cagr_start_year})')
    c = ws.cell(row=r, column=2, value=round(_cagr_start_rev, 1))
    _style(c, fill_hex=YELLOW, bold=False, h_align='right',
           number_format='#,##0.0')
    c.border = THIN_BOX
    cagr_start_row = r; r += 1

    _lbl(r, f'Ending Revenue ($M, {_cagr_end_label})')
    c = ws.cell(row=r, column=2, value=f'={_latest_hist_cl}{rev_row}')
    _style(c, fill_hex=LIGHT_GREEN, bold=False, h_align='right',
           number_format='#,##0.0')
    c.border = THIN_BOX
    cagr_end_row = r; r += 1

    _lbl(r, 'Number of Periods')
    c = ws.cell(row=r, column=2, value=_cagr_periods)
    _style(c, fill_hex=YELLOW, bold=False, h_align='right',
           number_format='0')
    c.border = THIN_BOX
    cagr_periods_row = r; r += 1

    _lbl(r, '5-Year Revenue CAGR', bold=True)
    c = ws.cell(row=r, column=2,
                value=f'=IF(B{cagr_start_row}>0,(B{cagr_end_row}/B{cagr_start_row})^(1/B{cagr_periods_row})-1,0)')
    _style(c, fill_hex=LIGHT_GREEN, bold=True, h_align='right',
           number_format=FMT_PCT)
    c.border = THIN_BOX
    cagr_row = r; r += 1

    # =========================================================================
    # SCENARIO ASSUMPTIONS TABLE  (at the bottom of the sheet)
    # Three cases (Best / Base / Weak) for each projection driver.
    # The dropdown cell in the title area selects which case to use.
    # =========================================================================
    r += 2
    _write_section_header(
        ws, r,
        'SCENARIO ASSUMPTIONS  (edit yellow cells to customize projections)',
        cols=total_cols)
    r += 1

    # Column headers for projection years + delta column
    delta_col = proj_start_col - 1               # column E (last hist col)
    delta_cl  = get_column_letter(delta_col)
    ws.cell(row=r, column=1, value='').fill = _fill(DARK_BLUE)
    for ci in range(2, delta_col):
        ws.cell(row=r, column=ci).fill = _fill(DARK_BLUE)
    c = ws.cell(row=r, column=delta_col, value='+/−')
    _style(c, fill_hex=DARK_BLUE, bold=True, font_color=WHITE, h_align='center')
    c.border = THIN_BOX
    for j, py in enumerate(proj_years):
        col = proj_start_col + j
        c = ws.cell(row=r, column=col, value=f'FY{py}E')
        _style(c, fill_hex=DARK_BLUE, bold=True, font_color=WHITE, h_align='center')
        c.border = THIN_BOX
    r += 1

    # Compute base-case values from historical averages
    def _hist_vals(num_data, denom_data, as_growth=False):
        """Compute historical ratios or growth rates from raw XBRL data.
        Excludes Q cumulative and annualized sentinel years."""
        vals = []
        _excl = set()
        if has_ann:
            _excl.add(_ltm_info.get('ltm_year'))
            _excl.add(ANN_YEAR)
        sorted_yrs = sorted(yr for yr in years if yr not in _excl)
        if as_growth:
            for idx in range(1, len(sorted_yrs)):
                y0, y1 = sorted_yrs[idx - 1], sorted_yrs[idx]
                v0 = num_data.get(y0)
                v1 = num_data.get(y1)
                if v0 and v1 and v0 > 0:
                    vals.append(v1 / v0 - 1.0)
        else:
            for yr in sorted_yrs:
                nv = num_data.get(yr)
                dv = denom_data.get(yr) if denom_data else None
                if nv is not None and dv and dv != 0:
                    vals.append(nv / dv)
        return _safe_avg(vals) if vals else None

    avg_rev_growth = 0.05  # Placeholder; base case uses CAGR formula reference
    avg_gp_margin  = _hist_vals(
        {yr: (inc['revenue'].get(yr) or 0) - (inc['cogs'].get(yr) or 0) for yr in years},
        inc['revenue']) or 0.40
    avg_rd_pct     = _hist_vals(inc['rd_expense'], inc['revenue']) or 0.05
    avg_sga_pct    = _hist_vals(inc['sga_expense'], inc['revenue']) or 0.10
    avg_tax_rate   = _hist_vals(inc['tax_expense'], inc['pretax_income']) or 0.21
    avg_ebitda_margin = _hist_vals(inc['ebitda'], inc['revenue']) or 0.30
    avg_da_pct     = _hist_vals(inc['da'], inc['revenue']) or 0.03
    avg_capex_pct  = _hist_vals(
        {yr: abs(cf['capex'].get(yr) or 0) for yr in years},
        inc['revenue']) or 0.05

    # Define metrics: (key, label, base_val, spread, invert)
    # invert=True means lower is "better" (costs / tax)
    metrics_def = [
        ('rev_growth',     'Revenue Growth (YoY %)',    avg_rev_growth,     0.03, False),
        ('gp_margin',      'Gross Profit Margin (%)',   avg_gp_margin,      0.03, False),
        ('rd_pct',         'R&D % of Sales',            avg_rd_pct,         0.02, True),
        ('sga_pct',        'SG&A % of Sales',           avg_sga_pct,        0.02, True),
        ('tax_rate',       'Tax Rate (%)',               avg_tax_rate,       0.03, True),
        ('ebitda_margin',  'EBITDA Margin (%)',          avg_ebitda_margin,  0.03, False),
        ('da_pct',         'D&A % of Revenue',           avg_da_pct,         0.01, True),
        ('capex_pct',      'Capex % of Revenue',         avg_capex_pct,      0.01, True),
    ]

    scenario_rows = {}  # {metric_key: {'best': row, 'base': row, 'weak': row}}
    FMT_DELTA = '+0.0%;-0.0%;0.0%'

    for metric_key, metric_label, base_val, spread, invert in metrics_def:
        # Sub-header for this metric
        _lbl(r, metric_label, bold=True, fill=LIGHT_BLUE)
        for ci in range(2, 2 + n + n_proj):
            ws.cell(row=r, column=ci).fill = _fill(LIGHT_BLUE)
        r += 1

        # Signed deltas: for "invert" metrics (costs), best = negative spread
        best_delta = -spread if invert else +spread
        weak_delta = +spread if invert else -spread

        # Row assignments (Best / Base / Weak in consecutive rows)
        best_row = r
        base_row = r + 1
        weak_row = r + 2

        scenario_rows[metric_key] = {
            'best': best_row,
            'base': base_row,
            'weak': weak_row,
        }

        # ── Best Case: delta cell + formula cells ──
        _lbl(best_row, '  Best Case', ind=1)
        c = ws.cell(row=best_row, column=delta_col, value=round(best_delta, 4))
        _style(c, fill_hex=YELLOW, bold=False, h_align='right',
               number_format=FMT_DELTA)
        c.border = THIN_BOX
        for j in range(n_proj):
            col = proj_start_col + j
            cl  = get_column_letter(col)
            c = ws.cell(row=best_row, column=col,
                        value=f'={cl}{base_row}+{delta_cl}{best_row}')
            _style(c, fill_hex=LIGHT_GREEN, bold=False, h_align='right',
                   number_format=FMT_PCT)
            c.border = THIN_BOX

        # ── Base Case ──
        _lbl(base_row, '  Base Case', ind=1)
        for j in range(n_proj):
            col = proj_start_col + j
            if metric_key == 'rev_growth':
                # Formula reference to 5-Year Revenue CAGR section
                c = ws.cell(row=base_row, column=col, value=f'=$B${cagr_row}')
                _style(c, fill_hex=LIGHT_GREEN, bold=False, h_align='right',
                       number_format=FMT_PCT)
            else:
                c = ws.cell(row=base_row, column=col, value=round(base_val, 4))
                _style(c, fill_hex=YELLOW, bold=False, h_align='right',
                       number_format=FMT_PCT)
            c.border = THIN_BOX

        # ── Weak Case: delta cell + formula cells ──
        _lbl(weak_row, '  Weak Case', ind=1)
        c = ws.cell(row=weak_row, column=delta_col, value=round(weak_delta, 4))
        _style(c, fill_hex=YELLOW, bold=False, h_align='right',
               number_format=FMT_DELTA)
        c.border = THIN_BOX
        for j in range(n_proj):
            col = proj_start_col + j
            cl  = get_column_letter(col)
            c = ws.cell(row=weak_row, column=col,
                        value=f'={cl}{base_row}+{delta_cl}{weak_row}')
            _style(c, fill_hex=LIGHT_GREEN, bold=False, h_align='right',
                   number_format=FMT_PCT)
            c.border = THIN_BOX

        r = weak_row + 1
        r += 1   # spacer between metrics

    # ── Deferred fill: assumptions band projection IF formulas ────────
    # Now that scenario table row numbers are known, fill projection
    # columns (F-J) of the assumptions band with IF(dropdown) formulas.
    asm_band_metrics = [
        (rev_growth_asm_row,     'rev_growth'),
        (gp_margin_asm_row,      'gp_margin'),
        (rd_pct_asm_row,         'rd_pct'),
        (sga_pct_asm_row,        'sga_pct'),
        (tax_rate_asm_row,       'tax_rate'),
        (ebitda_margin_asm_row,  'ebitda_margin'),
        (da_pct_asm_row,         'da_pct'),
        (capex_pct_asm_row,      'capex_pct'),
    ]
    for j in range(n_proj):
        col = proj_start_col + j
        col_letter = get_column_letter(col)
        for asm_row, metric_key in asm_band_metrics:
            best_ref = f'{col_letter}{scenario_rows[metric_key]["best"]}'
            base_ref = f'{col_letter}{scenario_rows[metric_key]["base"]}'
            weak_ref = f'{col_letter}{scenario_rows[metric_key]["weak"]}'
            formula = (
                f'=IF({dropdown_ref}="Best Case",{best_ref},'
                f'IF({dropdown_ref}="Weak Case",{weak_ref},{base_ref}))'
            )
            c = ws.cell(row=asm_row, column=col, value=formula)
            _style(c, fill_hex=LIGHT_GREEN, bold=False, h_align='right',
                   number_format=FMT_PCT)
            c.border = THIN_BOX

    # Return row map so DCF sheets can build cross-sheet references.
    # latest_col is the column letter for the most recent fiscal year.
    latest_col = get_column_letter(2 + n - 1)
    return {
        'latest_col':         latest_col,
        # Income Statement
        'revenue':            rev_row,
        'cogs':               cogs_row,
        'gp':                 gp_row,
        'ebit':               ebit_row,
        'da':                 da_row,
        'amortization':       amort_row,
        'transformation_costs': transform_row,
        'debt_extinguishment':  debt_ext_row,
        'ebitda':             ebitda_row,
        'sbc':                sbc_row,
        'adj_ebitda':         adj_ebitda_row,
        'interest_expense':   int_exp_row,
        'other_income':       other_inc_row,
        'pretax_income':      pretax_row,
        'tax_expense':        tax_row,
        'net_income':         ni_row,
        'eps_diluted':        eps_diluted_row,
        'shares_diluted':     shares_diluted_row,
        # Balance Sheet
        'cash':               cash_row,
        'st_investments':     st_inv_row,
        'accounts_rec':       ar_row,
        'inventory':          inventory_row,
        'other_current_a':    other_ca_row,
        'total_current_a':    total_ca_row,
        'accounts_pay':       ap_row,
        'accrued_liab':       accrued_row,
        'st_debt':            st_debt_row,
        'deferred_rev_cur':   deferred_rev_row,
        'other_current_l':    other_cl_row,
        'total_current_l':    total_cl_row,
        'lt_debt':            lt_debt_row,
        'total_assets':       total_assets_row,
        'total_liabilities':  total_liabilities_row,
        'total_equity':       total_equity_row,
        # Cash Flow
        'cf_net_income':      cf_ni_row,
        'cf_da':              cf_da_row,
        'operating_cf':       op_cf_row,
        'capex':              capex_row,
        'investing_cf':       inv_cf_row,
        'dividends':          dividends_row,
        'repurchases':        repurchases_row,
        'debt_issuance':      debt_iss_row,
        'debt_repayment':     debt_rep_row,
        'financing_cf':       fin_cf_row,
        'fx_effect':          fx_cf_row,
        'fcf':                fcf_row,
        'end_cash':           end_cf_row,
        # PP&E Schedule
        'ppe_beg':            ppe_beg_sch_row,
        'ppe_capex':          ppe_capex_sch_row,
        'ppe_depr':           ppe_depr_sch_row,
        'ppe_end':            ppe_end_sch_row,
        'da_capex_pct':       da_capex_pct_row,
        # Assumptions band rows (scenario-driven, for DCF cross-sheet refs)
        'rev_growth_asm':     rev_growth_asm_row,
        'ebitda_margin_asm':  ebitda_margin_asm_row,
        'da_pct_asm':         da_pct_asm_row,
        'capex_pct_asm':      capex_pct_asm_row,
        'proj_start_col':     proj_start_col,
    }


# ─────────────────────────────────────────────────────────────────────────────
# SHEET 2 – DCF MODEL
# ─────────────────────────────────────────────────────────────────────────────

