"""DCF model sheet builder."""

from typing import Dict
from .shared import *

def _calc_dcf_assumptions(fd: Dict) -> Dict:
    """Derive projection assumptions from historical data."""
    years_desc = fd['years']
    # Use only actual full-year annual data for averages
    _ltm = fd.get('ltm_info', {})
    _ann_yrs = {_ltm.get('ltm_year')} if _ltm else set()
    # Also exclude any float sentinel (ANN_YEAR)
    hist_yrs = [yr for yr in years_desc if isinstance(yr, int) and yr not in _ann_yrs]
    latest_yr  = hist_yrs[0] if hist_yrs else years_desc[0]
    inc  = fd['income_statement']
    cf   = fd['cash_flow']
    bs   = fd['balance_sheet']

    def v(d, yr): return _val(d, yr)

    # Revenue CAGR over available history
    rev_growths = []
    for i in range(len(hist_yrs) - 1):
        r0 = v(inc['revenue'], hist_yrs[i+1])
        r1 = v(inc['revenue'], hist_yrs[i])
        if r0 and r1 and r0 > 0:
            rev_growths.append((r1 / r0) - 1.0)
    avg_rev_growth = _safe_avg(rev_growths) if rev_growths else 0.05

    # EBITDA margin
    ebitda_margins = []
    for yr in hist_yrs:
        rev = v(inc['revenue'], yr)
        ebt = v(inc['ebitda'], yr)
        if rev and ebt:
            ebitda_margins.append(ebt / rev)
    avg_ebitda_margin = _safe_avg(ebitda_margins) if ebitda_margins else 0.20

    # D&A % revenue
    da_pcts = []
    for yr in hist_yrs:
        rev = v(inc['revenue'], yr)
        da  = v(inc['da'], yr)
        if rev and da:
            da_pcts.append(da / rev)
    avg_da_pct = _safe_avg(da_pcts) if da_pcts else 0.05

    # Capex % revenue
    capex_pcts = []
    for yr in hist_yrs:
        rev  = v(inc['revenue'], yr)
        capx = v(cf['capex'], yr)
        if rev and capx:
            capex_pcts.append(abs(capx) / rev)
    avg_capex_pct = _safe_avg(capex_pcts) if capex_pcts else 0.04

    # Effective tax rate
    tax_rates = []
    for yr in hist_yrs:
        pre = v(inc['pretax_income'], yr)
        tax = v(inc['tax_expense'], yr)
        if pre and tax and pre > 0:
            tax_rates.append(tax / pre)
    avg_tax = _safe_avg(tax_rates) if tax_rates else 0.21

    # Implied cost of debt: interest expense / average total debt
    cod_estimates = []
    for yr in hist_yrs:
        int_exp   = abs(v(inc['interest_expense'], yr) or 0)
        total_dbt = (v(bs['lt_debt'], yr) or 0) + (v(bs['st_debt'], yr) or 0)
        if int_exp and total_dbt > 0:
            cod_estimates.append(int_exp / total_dbt)
    avg_cod = round(_safe_avg(cod_estimates), 4) if cod_estimates else 0.05

    # Capital structure inputs — v() already returns $M via _val(scale=1e6)
    st_debt  = (v(bs['st_debt'],       latest_yr) or 0)
    lt_debt  = (v(bs['lt_debt'],       latest_yr) or 0)
    cash     = (v(bs['cash'],          latest_yr) or 0)
    bk_eq    = (v(bs['total_equity'],  latest_yr) or 0)
    net_debt = round(st_debt + lt_debt - cash, 1)

    # Market cap: use yfinance-derived value if available, else fall back to book equity
    actual_mc = (fd.get('market_cap') or {}).get(latest_yr)
    if actual_mc is not None:
        market_cap = round(actual_mc / 1e6, 1)
    else:
        market_cap = round(bk_eq, 1) if bk_eq else 0.0

    return {
        # CAPM inputs
        'rf_rate':          0.045,          # 10Y US Treasury yield (editable)
        'erp':              0.05,           # Equity risk premium   (editable)
        'beta':             1.0,            # Beta vs market        (editable)
        'cost_of_debt':     avg_cod,        # Implied from interest expense / debt
        # Capital structure
        'market_cap':       market_cap,     # From yfinance (or book equity fallback)
        'net_debt':         net_debt,       # (ST + LT Debt) - Cash  ($M)
        # Projection assumptions
        'terminal_growth':  0.025,
        'rev_growth':       round(avg_rev_growth, 4),
        'ebitda_margin':    round(avg_ebitda_margin, 4),
        'da_pct':           round(avg_da_pct, 4),
        'capex_pct':        round(avg_capex_pct, 4),
        'tax_rate':         round(avg_tax, 4),
    }


# ─────────────────────────────────────────────────────────────────────────────
# SHEET 2 – WACC  (Weighted Average Cost of Capital)
# ─────────────────────────────────────────────────────────────────────────────


def _write_dcf_model(ws, company_info: Dict, fd: Dict, fs_rows: Dict,
                     wacc_rows: Dict = None):
    years_desc = fd['years']
    years      = list(reversed(years_desc))    # oldest → newest
    latest_yr  = years[-1]
    asm  = _calc_dcf_assumptions(fd)
    proj_years = [latest_yr + i for i in range(1, 6)]   # 5-year projection

    # Cross-sheet reference helpers — link back to the Financial Statements sheet
    FS  = "'Financial Statements'"
    fsc = fs_rows['latest_col']   # column letter for latest fiscal year on FS
    def _fs(key):
        """Formula referencing a cell on the Financial Statements sheet."""
        return f"={FS}!{fsc}{fs_rows[key]}"

    # Column map: A=label, B=base year, C-G=proj years 1-5
    _set_col_widths(ws, {1: 38, 2: 16, 3: 14, 4: 14, 5: 14, 6: 14, 7: 14})

    r = 1
    title = ws.cell(row=r, column=1,
                    value=f"{company_info['name']}  ({company_info['ticker']}) — DCF Valuation  ($ in millions)")
    _style(title, fill_hex=DARK_BLUE, bold=True, font_color=WHITE)
    ws.merge_cells(start_row=r, start_column=1, end_row=r, end_column=7)
    ws.row_dimensions[r].height = 18
    r += 1
    sub = ws.cell(row=r, column=1,
                  value=f"Base Year: FY{latest_yr}  |  Yellow cells = editable inputs  |  All amounts in $M")
    _style(sub, fill_hex=XLIGHT_BLUE, italic=True)
    ws.merge_cells(start_row=r, start_column=1, end_row=r, end_column=7)
    r += 2

    # ── Assumptions ───────────────────────────────────────────────────
    _write_section_header(ws, r, 'KEY ASSUMPTIONS  (edit yellow cells)', cols=4); r += 1

    def _asm_row(row, label, value, fmt=FMT_PCT, cell_ref=None):
        """Editable yellow input cell."""
        lbl = ws.cell(row=row, column=1, value=label)
        _style(lbl, bold=False)
        c = ws.cell(row=row, column=2, value=value)
        _style(c, fill_hex=YELLOW, bold=True, h_align='right', number_format=fmt)
        c.border = THIN_BOX
        return f'$B${row}'

    def _fml_row(row, label, value_or_formula, fmt=FMT_PCT, bold=False):
        """Non-editable formula / derived cell (light blue)."""
        lbl = ws.cell(row=row, column=1, value=label)
        _style(lbl, bold=False)
        c = ws.cell(row=row, column=2, value=value_or_formula)
        _style(c, fill_hex=XLIGHT_BLUE, bold=bold, h_align='right', number_format=fmt)
        c.border = THIN_BOX
        return f'$B${row}'

    # ── WACC & Tax Rate (from WACC sheet) ────────────────────────────────────
    WC = "'WACC'"
    if wacc_rows:
        wacc_ref = _fml_row(r, 'WACC  (from WACC sheet)',
                            f"={WC}!$B${wacc_rows['wacc']}", bold=True);       r += 1
        tax_ref  = _fml_row(r, 'Tax Rate  (from WACC sheet)',
                            f"={WC}!$B${wacc_rows['tax_rate']}");              r += 1
    else:
        # Fallback: standalone WACC (for backward compat)
        rf_ref      = _asm_row(r, 'Risk-Free Rate', asm['rf_rate']);           r += 1
        erp_ref     = _asm_row(r, 'Equity Risk Premium', asm['erp']);            r += 1
        beta_ref    = _asm_row(r, 'Beta', asm['beta'], fmt='0.00');            r += 1
        coe_ref     = _fml_row(r, 'Cost of Equity',
                                f'={rf_ref}+{beta_ref}*{erp_ref}');            r += 1
        cod_ref     = _asm_row(r, 'Cost of Debt', asm['cost_of_debt']);        r += 1
        mktcap_ref  = _asm_row(r, 'Market Cap ($M)', asm['market_cap'],
                                fmt=FMT_DOLLAR);                               r += 1
        netdebt_ref = _fml_row(r, 'Net Debt ($M)',
            f"={FS}!{fsc}{fs_rows['st_debt']}+{FS}!{fsc}{fs_rows['lt_debt']}"
            f"-{FS}!{fsc}{fs_rows['cash']}", fmt=FMT_DOLLAR);                 r += 1
        _denom     = f'({mktcap_ref}+{netdebt_ref})'
        eq_wt_ref  = _fml_row(r, 'Equity Weight',
                                f'=IF({_denom}<>0,{mktcap_ref}/{_denom},1)');  r += 1
        dbt_wt_ref = _fml_row(r, 'Debt Weight',
                                f'=IF({_denom}<>0,{netdebt_ref}/{_denom},0)'); r += 1
        tax_ref    = _asm_row(r, 'Tax Rate', asm['tax_rate']);                 r += 1
        wacc_ref   = _fml_row(r, 'WACC',
                                f'={eq_wt_ref}*{coe_ref}+{dbt_wt_ref}*{cod_ref}*(1-{tax_ref})',
                                bold=True);                                     r += 1
    r += 1

    # ── Projection Assumptions ────────────────────────────────────────────────
    tg_ref          = _asm_row(r, 'Terminal Growth Rate',           asm['terminal_growth']);     r += 1

    # 4 metrics driven by FS scenario dropdown (read-only display of base-case)
    fs_psc = fs_rows['proj_start_col']  # FS projection start column (6 = col F)
    fs_proj_cl = [get_column_letter(fs_psc + j) for j in range(5)]  # F, G, H, I, J

    def _fs_asm_ref(i, fs_asm_key):
        """Cross-sheet reference to FS assumptions band for projection year i."""
        return f"{FS}!{fs_proj_cl[i]}{fs_rows[fs_asm_key]}"

    _fml_row(r, 'Revenue Growth Rate  (from FS scenario)',
             f"={_fs_asm_ref(0, 'rev_growth_asm')}");  r += 1
    _fml_row(r, 'EBITDA Margin  (from FS scenario)',
             f"={_fs_asm_ref(0, 'ebitda_margin_asm')}"); r += 1
    _fml_row(r, 'D&A % of Revenue  (from FS scenario)',
             f"={_fs_asm_ref(0, 'da_pct_asm')}");       r += 1
    _fml_row(r, 'Capex % of Revenue  (from FS scenario)',
             f"={_fs_asm_ref(0, 'capex_pct_asm')}");    r += 1
    r += 1

    # ── Historical Base Data ──────────────────────────────────────────
    _write_section_header(ws, r, f'HISTORICAL BASE  (FY{latest_yr})', cols=3); r += 1

    base_data = [
        ('Revenue ($M)',          _fs('revenue'),                                 FMT_DOLLAR),
        ('EBITDA ($M)',           _fs('ebitda'),                                  FMT_DOLLAR),
        ('EBIT ($M)',             _fs('ebit'),                                    FMT_DOLLAR),
        ('D&A ($M)',              _fs('da'),                                      FMT_DOLLAR),
        ('Net Income ($M)',       _fs('net_income'),                              FMT_DOLLAR),
        ('Capex ($M)',            f"=-{FS}!{fsc}{fs_rows['capex']}",              FMT_DOLLAR),
        ('Free Cash Flow ($M)',   _fs('fcf'),                                     FMT_DOLLAR),
        ('Short-term Debt ($M)',  _fs('st_debt'),                                 FMT_DOLLAR),
        ('Long-term Debt ($M)',   _fs('lt_debt'),                                 FMT_DOLLAR),
        ('Cash ($M)',             _fs('cash'),                                    FMT_DOLLAR),
        ('Net Debt ($M)',         0,                                              FMT_DOLLAR),
        ('Diluted Shares (M)',    _fs('shares_diluted'),                          FMT_DOLLAR),
    ]
    # Store cell refs for valuation section
    base_rows = {}
    for label, value, fmt in base_data:
        lbl = ws.cell(row=r, column=1, value=f'  {label}')
        _style(lbl)
        c = ws.cell(row=r, column=2, value=value)
        _style(c, fill_hex=LIGHT_GRAY, h_align='right', number_format=fmt)
        base_rows[label] = r
        r += 1
    # Net Debt = Short-term Debt + Long-term Debt - Cash  [formula]
    ws.cell(row=base_rows['Net Debt ($M)'], column=2,
            value=f'=B{base_rows["Short-term Debt ($M)"]}+B{base_rows["Long-term Debt ($M)"]}-B{base_rows["Cash ($M)"]}')
    r += 1

    # ── 5-Year Projections ────────────────────────────────────────────
    _write_section_header(ws, r, '5-YEAR FREE CASH FLOW PROJECTIONS', cols=7); r += 1

    # Column headers
    ws.cell(row=r, column=1, value='').fill = _fill(DARK_BLUE)
    for i, py in enumerate(proj_years):
        c = ws.cell(row=r, column=3 + i, value=f'FY{py}E')
        _style(c, fill_hex=DARK_BLUE, bold=True, font_color=WHITE, h_align='center')
        c.border = THIN_BOX
    ws.cell(row=r, column=2, value=f'FY{latest_yr} (Base)')
    _style(ws.cell(row=r, column=2), fill_hex=MED_BLUE, bold=True, font_color=WHITE,
           h_align='center')
    r += 1

    proj_start_row = r

    # Revenue projections
    rev_row = r
    ws.cell(row=r, column=1, value='Revenue ($M)')
    # Base year — references Financial Statements sheet
    c = ws.cell(row=r, column=2, value=_fs('revenue'))
    _style(c, fill_hex=LIGHT_GRAY, h_align='right', number_format=FMT_DOLLAR)

    for i in range(5):
        col = 3 + i
        prev_col = get_column_letter(col - 1)
        formula = f'={prev_col}{r}*(1+{_fs_asm_ref(i, "rev_growth_asm")})'
        c = ws.cell(row=r, column=col, value=formula)
        _style(c, h_align='right', number_format=FMT_DOLLAR)
    r += 1

    # EBITDA
    ebitda_row = r
    ws.cell(row=r, column=1, value='EBITDA ($M)')
    c = ws.cell(row=r, column=2, value=_fs('ebitda'))
    _style(c, fill_hex=LIGHT_GRAY, h_align='right', number_format=FMT_DOLLAR)
    for i in range(5):
        col = 3 + i
        rev_cell = f'{get_column_letter(col)}{rev_row}'
        c = ws.cell(row=r, column=col, value=f'={rev_cell}*{_fs_asm_ref(i, "ebitda_margin_asm")}')
        _style(c, h_align='right', number_format=FMT_DOLLAR)
    r += 1

    # D&A
    da_row = r
    ws.cell(row=r, column=1, value='D&A ($M)')
    c = ws.cell(row=r, column=2, value=_fs('da'))
    _style(c, fill_hex=LIGHT_GRAY, h_align='right', number_format=FMT_DOLLAR)
    for i in range(5):
        col = 3 + i
        rev_cell = f'{get_column_letter(col)}{rev_row}'
        c = ws.cell(row=r, column=col, value=f'={rev_cell}*{_fs_asm_ref(i, "da_pct_asm")}')
        _style(c, h_align='right', number_format=FMT_DOLLAR)
    r += 1

    # EBIT
    ebit_row = r
    ws.cell(row=r, column=1, value='EBIT ($M)')
    c = ws.cell(row=r, column=2, value=_fs('ebit'))
    _style(c, fill_hex=LIGHT_GRAY, h_align='right', number_format=FMT_DOLLAR)
    for i in range(5):
        col = 3 + i
        cl = get_column_letter(col)
        c = ws.cell(row=r, column=col, value=f'={cl}{ebitda_row}-{cl}{da_row}')
        _style(c, h_align='right', number_format=FMT_DOLLAR)
    r += 1

    # NOPAT = EBIT * (1 - tax)
    nopat_row = r
    ws.cell(row=r, column=1, value='NOPAT  [EBIT × (1 − Tax)]')
    for i in range(5):
        col = 3 + i
        cl  = get_column_letter(col)
        c   = ws.cell(row=r, column=col, value=f'={cl}{ebit_row}*(1-{tax_ref})')
        _style(c, h_align='right', number_format=FMT_DOLLAR)
    r += 1

    # Capex
    capex_row = r
    ws.cell(row=r, column=1, value='Capex ($M)')
    c = ws.cell(row=r, column=2, value=f"=-{FS}!{fsc}{fs_rows['capex']}")
    _style(c, fill_hex=LIGHT_GRAY, h_align='right', number_format=FMT_DOLLAR)
    for i in range(5):
        col = 3 + i
        rev_cell = f'{get_column_letter(col)}{rev_row}'
        c = ws.cell(row=r, column=col, value=f'={rev_cell}*{_fs_asm_ref(i, "capex_pct_asm")}')
        _style(c, h_align='right', number_format=FMT_DOLLAR)
    r += 1

    # Unlevered FCF = NOPAT + D&A - Capex
    fcf_row = r
    ws.cell(row=r, column=1, value='Unlevered Free Cash Flow ($M)')
    _style(ws.cell(row=r, column=1), bold=True)
    c = ws.cell(row=r, column=2, value=_fs('fcf'))
    _style(c, fill_hex=LIGHT_GREEN, bold=True, h_align='right', number_format=FMT_DOLLAR)
    for i in range(5):
        col = 3 + i
        cl  = get_column_letter(col)
        c   = ws.cell(row=r, column=col,
                      value=f'={cl}{nopat_row}+{cl}{da_row}-{cl}{capex_row}')
        _style(c, fill_hex=LIGHT_GREEN, bold=True, h_align='right', number_format=FMT_DOLLAR)
    r += 1

    # Discount factor  1/(1+WACC)^n
    disc_row = r
    ws.cell(row=r, column=1, value='Discount Factor  [1/(1+WACC)ⁿ]')
    for i in range(5):
        col = 3 + i
        n   = i + 1
        c   = ws.cell(row=r, column=col, value=f'=1/(1+{wacc_ref})^{n}')
        _style(c, fill_hex=LIGHT_GRAY, h_align='right', number_format='0.0000')
    r += 1

    # PV of FCF
    pv_fcf_row = r
    ws.cell(row=r, column=1, value='PV of FCF ($M)')
    _style(ws.cell(row=r, column=1), bold=True)
    for i in range(5):
        col = 3 + i
        cl  = get_column_letter(col)
        c   = ws.cell(row=r, column=col, value=f'={cl}{fcf_row}*{cl}{disc_row}')
        _style(c, bold=True, h_align='right', number_format=FMT_DOLLAR)
    r += 2

    # ── Valuation ─────────────────────────────────────────────────────
    _write_section_header(ws, r, 'VALUATION SUMMARY', cols=4); r += 1

    sum_pv_formula = '=' + '+'.join(f'{get_column_letter(3+i)}{pv_fcf_row}' for i in range(5))
    # Terminal FCF = last proj year FCF * (1 + terminal growth)
    last_fcf_cell = f'{get_column_letter(7)}{fcf_row}'
    tv_formula   = f'={last_fcf_cell}*(1+{tg_ref})/({wacc_ref}-{tg_ref})'
    pv_tv_formula = f'={get_column_letter(7)}{disc_row}*C{r+3}'   # will set row dynamically

    val_rows = {}

    def _val_row(label, formula, fmt=FMT_DOLLAR, bold=False, fill=None):
        nonlocal r
        lbl = ws.cell(row=r, column=1, value=label)
        _style(lbl, bold=bold)
        c = ws.cell(row=r, column=2, value=formula)
        _style(c, fill_hex=fill, bold=bold, h_align='right', number_format=fmt)
        if bold and fill:
            c.border = BOT_MED
        ref = f'B{r}'
        val_rows[label] = r
        r += 1
        return ref

    sum_pv_ref = _val_row('Sum of PV (FCF)', sum_pv_formula)
    tv_ref     = _val_row('Terminal Value ($M)', tv_formula)
    pv_tv      = _val_row('PV of Terminal Value ($M)',
                           f'={get_column_letter(7)}{disc_row}*{tv_ref}')
    ev_ref     = _val_row('Enterprise Value ($M)',
                           f'={sum_pv_ref}+B{val_rows["PV of Terminal Value ($M)"]}',
                           bold=True, fill=LIGHT_BLUE)
    # Net debt from base data
    nd_row_num = base_rows.get('Net Debt ($M)', r)
    nd_ref     = f'B{nd_row_num}'
    eq_ref     = _val_row('Less: Net Debt ($M)', f'={nd_ref}')
    eq_val_ref = _val_row('Equity Value ($M)',
                           f'={ev_ref}-{eq_ref}', bold=True, fill=LIGHT_GREEN)
    if wacc_rows:
        sh_ref = f"={WC}!$B${wacc_rows['diluted_shares']}"
        _val_row('Diluted Shares Outstanding (M)', sh_ref)
        sh_cell = f"B{val_rows['Diluted Shares Outstanding (M)']}"
    else:
        sh_row_num = base_rows.get('Diluted Shares (M)', r)
        sh_ref     = f'B{sh_row_num}'
        _val_row('Diluted Shares Outstanding (M)', f'={sh_ref}')
        sh_cell = f'B{val_rows["Diluted Shares Outstanding (M)"]}'
    _val_row('Implied Share Price ($)',
             f'={eq_val_ref}/{sh_cell}*1',
             fmt='$#,##0.00', bold=True, fill=LIGHT_GREEN)

    ws.freeze_panes = 'A3'


# ─────────────────────────────────────────────────────────────────────────────
# SHEET 3 – DATA VALIDATION
# ─────────────────────────────────────────────────────────────────────────────

# Extra palette entries for validation results
DARK_RED   = "C00000"
PASS_GREEN = "548235"
FAIL_RED   = "FFE2CC"      # same as LIGHT_RED


