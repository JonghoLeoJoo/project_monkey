"""WACC sheet builder."""

from typing import Dict
from .shared import *

def _write_wacc_sheet(ws, company_info: Dict, fd: Dict, fs_rows: Dict):
    """Build the WACC analysis sheet.  Returns *wacc_rows* dict for DCF."""
    FS  = "'Financial Statements'"
    fsc = fs_rows['latest_col']

    wacc_inputs   = fd.get('wacc_inputs', {})
    current_price = wacc_inputs.get('current_price', {})
    shares_bk     = wacc_inputs.get('shares_breakdown', {})
    comparables   = wacc_inputs.get('comparables', [])

    _set_col_widths(ws, {1: 34, 2: 16, 3: 16, 4: 16, 5: 16, 6: 16, 7: 14, 8: 16})
    total_cols = 8

    r = 1
    wacc_rows = {}   # row numbers to export

    # ── Title ────────────────────────────────────────────────────────
    title = ws.cell(row=r, column=1,
                    value=f"{company_info['name']}  ({company_info['ticker']}) "
                          f"— WACC Analysis  ($ in millions)")
    _style(title, fill_hex=DARK_BLUE, bold=True, font_color=WHITE)
    ws.merge_cells(start_row=r, start_column=1, end_row=r, end_column=total_cols)
    ws.row_dimensions[r].height = 18
    r += 1
    sub = ws.cell(row=r, column=1,
                  value="Yellow cells = editable inputs  |  "
                        "All amounts in $M, shares in millions")
    _style(sub, fill_hex=XLIGHT_BLUE, italic=True)
    ws.merge_cells(start_row=r, start_column=1, end_row=r, end_column=total_cols)
    r += 2

    # ── Local helpers (same pattern as DCF) ──────────────────────────
    def _asm_row(row, label, value, fmt=FMT_PCT):
        lbl = ws.cell(row=row, column=1, value=label)
        _style(lbl)
        c = ws.cell(row=row, column=2, value=value)
        _style(c, fill_hex=YELLOW, bold=True, h_align='right', number_format=fmt)
        c.border = THIN_BOX
        return f'$B${row}'

    def _fml_row(row, label, formula, fmt=FMT_PCT, bold=False, fill=XLIGHT_BLUE):
        lbl = ws.cell(row=row, column=1, value=label)
        _style(lbl, bold=bold)
        c = ws.cell(row=row, column=2, value=formula)
        _style(c, fill_hex=fill, bold=bold, h_align='right', number_format=fmt)
        c.border = THIN_BOX
        return f'$B${row}'

    # ══════════════════════════════════════════════════════════════════
    # SECTION 1 – STOCK PRICE & DATE
    # ══════════════════════════════════════════════════════════════════
    _write_section_header(ws, r, 'STOCK PRICE & DATE', cols=2); r += 1

    price_date = current_price.get('date') or ''
    price_val  = current_price.get('price') or 0
    _asm_row(r, 'Price Date', price_date, fmt='@'); r += 1

    wacc_rows['price'] = r
    price_ref = _asm_row(r, 'Share Price ($)', price_val, fmt=FMT_DOLLAR2); r += 1
    r += 1

    # ══════════════════════════════════════════════════════════════════
    # SECTION 2 – DILUTED SHARES OUTSTANDING
    # ══════════════════════════════════════════════════════════════════
    _write_section_header(ws, r, 'DILUTED SHARES OUTSTANDING  (from 10-K, in millions)',
                          cols=2); r += 1

    basic_ref  = _asm_row(r, 'Basic Shares Outstanding (M)',
                           shares_bk.get('basic', 0), fmt=FMT_DOLLAR); r += 1
    rsus_ref   = _asm_row(r, '(+) Restricted Stock / RSUs (M)',
                           shares_bk.get('rsus', 0), fmt=FMT_DOLLAR); r += 1
    opts_ref   = _asm_row(r, '(+) Options & Warrants (ITM) (M)',
                           shares_bk.get('options', 0), fmt=FMT_DOLLAR); r += 1
    conv_d_ref = _asm_row(r, '(+) Convertible Debt (ITM) (M)',
                           shares_bk.get('conv_debt', 0), fmt=FMT_DOLLAR); r += 1
    conv_p_ref = _asm_row(r, '(+) Convertible Preferred (ITM) (M)',
                           shares_bk.get('conv_pref', 0), fmt=FMT_DOLLAR); r += 1

    wacc_rows['diluted_shares'] = r
    diluted_ref = _fml_row(r, 'Net Diluted Shares Outstanding (M)',
                           f'={basic_ref}+{rsus_ref}+{opts_ref}+{conv_d_ref}+{conv_p_ref}',
                           fmt=FMT_DOLLAR, bold=True, fill=LIGHT_GREEN); r += 1
    r += 1

    # ══════════════════════════════════════════════════════════════════
    # SECTION 3 – COST OF DEBT
    # ══════════════════════════════════════════════════════════════════
    _write_section_header(ws, r, 'COST OF DEBT', cols=2); r += 1

    wacc_rows['cod'] = r
    cod_ref = _asm_row(r, 'Cost of Debt (YTM)',
                       wacc_inputs.get('implied_cod', 0.05)); r += 1

    wacc_rows['tax_rate'] = r
    tax_ref = _fml_row(r, 'Tax Rate  (final year actual)',
                       f"=IF({FS}!{fsc}{fs_rows['pretax_income']}>0,"
                       f"{FS}!{fsc}{fs_rows['tax_expense']}"
                       f"/{FS}!{fsc}{fs_rows['pretax_income']},"
                       f"0.21)"); r += 1

    wacc_rows['cod_at'] = r
    cod_at_ref = _fml_row(r, 'Cost of Debt (After Tax)',
                          f'={cod_ref}*(1-{tax_ref})',
                          bold=True); r += 1
    r += 1

    # ══════════════════════════════════════════════════════════════════
    # SECTION 4 – COST OF EQUITY  (CAPM with Comparable Companies)
    # ══════════════════════════════════════════════════════════════════
    _write_section_header(ws, r, 'COST OF EQUITY  (CAPM with Comparable Companies)',
                          cols=total_cols); r += 1

    wacc_rows['rf_rate'] = r
    rf_ref = _asm_row(r, 'Risk-Free Rate  (CNBC US10Y)',
                      wacc_inputs.get('treasury_yield', 0.045)); r += 1
    wacc_rows['erp'] = r
    erp_ref = _asm_row(r, 'Equity Risk Premium  (Kroll ERP)',
                       wacc_inputs.get('kroll_erp', 0.05)); r += 1
    r += 1

    # ── Comparable Companies Table ───────────────────────────────────
    comp_hdr = ws.cell(row=r, column=1, value='COMPARABLE COMPANIES')
    _style(comp_hdr, fill_hex=MED_BLUE, bold=True, font_color=WHITE)
    ws.merge_cells(start_row=r, start_column=1, end_row=r, end_column=total_cols)
    r += 1

    # Column headers
    for ci, hdr in enumerate([
        'Company Name', 'Observed Beta', 'Share Price ($)',
        'Diluted Shares (M)', 'Market Cap ($M)', 'Net Debt ($M)',
        'Tax Rate', 'De-levered Beta',
    ]):
        c = ws.cell(row=r, column=1 + ci, value=hdr)
        _style(c, fill_hex=DARK_BLUE, bold=True, font_color=WHITE, h_align='center')
        c.border = THIN_BOX
    r += 1

    # 10 comparable company rows
    comp_start_row = r
    for idx in range(10):
        comp = comparables[idx] if idx < len(comparables) else None

        # Col A – Company Name
        nc = ws.cell(row=r, column=1, value=comp['name'] if comp else '')
        _style(nc, fill_hex=YELLOW, h_align='left')
        nc.border = THIN_BOX

        # Col B – Observed Beta
        bc = ws.cell(row=r, column=2, value=comp['beta'] if comp else '')
        _style(bc, fill_hex=YELLOW, bold=True, h_align='right', number_format='0.00')
        bc.border = THIN_BOX

        # Col C – Share Price ($)
        pc = ws.cell(row=r, column=3, value=comp['price'] if comp else '')
        _style(pc, fill_hex=YELLOW, bold=True, h_align='right', number_format=FMT_DOLLAR2)
        pc.border = THIN_BOX

        # Col D – Diluted Shares (M)
        sc = ws.cell(row=r, column=4,
                     value=round(comp['shares'] / 1e6, 1) if comp else '')
        _style(sc, fill_hex=YELLOW, bold=True, h_align='right', number_format=FMT_DOLLAR)
        sc.border = THIN_BOX

        # Col E – Market Cap ($M) = Price × Shares (formula)
        mc = ws.cell(row=r, column=5, value=f'=C{r}*D{r}')
        _style(mc, fill_hex=XLIGHT_BLUE, h_align='right', number_format=FMT_DOLLAR)
        mc.border = THIN_BOX

        # Col F – Net Debt ($M)
        nd = ws.cell(row=r, column=6,
                     value=round(comp['net_debt'] / 1e6, 1) if comp else '')
        _style(nd, fill_hex=YELLOW, bold=True, h_align='right', number_format=FMT_DOLLAR)
        nd.border = THIN_BOX

        # Col G – Tax Rate
        tr = ws.cell(row=r, column=7, value=comp['tax_rate'] if comp else '')
        _style(tr, fill_hex=YELLOW, bold=True, h_align='right', number_format=FMT_PCT)
        tr.border = THIN_BOX

        # Col H – De-levered Beta  = Beta / (1 + (1-Tax) × NetDebt/MktCap)
        dlb = ws.cell(row=r, column=8,
                      value=f'=IF(OR(B{r}="",E{r}=0),"",B{r}/(1+(1-G{r})*F{r}/E{r}))')
        _style(dlb, fill_hex=XLIGHT_BLUE, h_align='right', number_format='0.0000')
        dlb.border = THIN_BOX

        r += 1

    comp_end_row = r - 1
    r += 1  # blank row

    # ── Beta Derivation ──────────────────────────────────────────────
    h_range = f'H{comp_start_row}:H{comp_end_row}'
    wacc_rows['avg_beta'] = r
    # If 3+ comps: trim highest & lowest outlier, average the rest.
    # If 1-2 comps: plain average.  If 0: fallback to 1.0.
    avg_beta_ref = _fml_row(
        r, 'Industry Avg De-levered Beta  (excl. high/low)',
        f'=IF(COUNT({h_range})>2,'
        f'(SUM({h_range})-MAX({h_range})-MIN({h_range}))/(COUNT({h_range})-2),'
        f'IF(COUNT({h_range})>0,AVERAGE({h_range}),1))',
        fmt='0.0000'); r += 1
    r += 1

    # Target company inputs
    wacc_rows['tgt_net_debt'] = r
    tgt_nd_ref = _fml_row(r, 'Target Company Net Debt ($M)',
                          f"={FS}!{fsc}{fs_rows['st_debt']}"
                          f"+{FS}!{fsc}{fs_rows['lt_debt']}"
                          f"-{FS}!{fsc}{fs_rows['cash']}",
                          fmt=FMT_DOLLAR); r += 1

    wacc_rows['tgt_mktcap'] = r
    tgt_mc_ref = _fml_row(r, 'Target Company Market Cap ($M)',
                          f'={price_ref}*{diluted_ref}',
                          fmt=FMT_DOLLAR); r += 1

    wacc_rows['beta'] = r
    relevered_ref = _fml_row(
        r, 'Target Re-levered Beta',
        f'=IF({tgt_mc_ref}=0,{avg_beta_ref},'
        f'{avg_beta_ref}*(1+(1-{tax_ref})*{tgt_nd_ref}/{tgt_mc_ref}))',
        fmt='0.0000'); r += 1

    wacc_rows['coe'] = r
    coe_ref = _fml_row(r, 'Cost of Equity  [Rf + Beta × ERP]',
                       f'={rf_ref}+{relevered_ref}*{erp_ref}',
                       bold=True, fill=LIGHT_BLUE); r += 1
    r += 1

    # ══════════════════════════════════════════════════════════════════
    # SECTION 5 – WACC CALCULATION
    # ══════════════════════════════════════════════════════════════════
    _write_section_header(ws, r, 'WACC CALCULATION', cols=2); r += 1

    wacc_rows['mktcap'] = r
    wacc_mc_ref = _fml_row(r, 'Market Cap ($M)',
                           f'={price_ref}*{diluted_ref}',
                           fmt=FMT_DOLLAR); r += 1

    wacc_rows['net_debt'] = r
    wacc_nd_ref = _fml_row(r, 'Net Debt ($M)',
                           f"={FS}!{fsc}{fs_rows['st_debt']}"
                           f"+{FS}!{fsc}{fs_rows['lt_debt']}"
                           f"-{FS}!{fsc}{fs_rows['cash']}",
                           fmt=FMT_DOLLAR); r += 1

    wacc_rows['ev'] = r
    ev_ref = _fml_row(r, 'Enterprise Value ($M)',
                      f'={wacc_mc_ref}+{wacc_nd_ref}',
                      fmt=FMT_DOLLAR, bold=True); r += 1

    wacc_rows['eq_weight'] = r
    eq_wt_ref = _fml_row(r, 'Equity Weight',
                         f'=IF({ev_ref}<>0,{wacc_mc_ref}/{ev_ref},1)'); r += 1

    wacc_rows['debt_weight'] = r
    debt_wt_ref = _fml_row(r, 'Debt Weight',
                           f'=IF({ev_ref}<>0,{wacc_nd_ref}/{ev_ref},0)'); r += 1

    # WACC = Equity Weight × CoE + Debt Weight × CoD(AT)
    wacc_rows['wacc'] = r
    wacc_lbl = ws.cell(row=r, column=1, value='WACC')
    _style(wacc_lbl, bold=True)
    wacc_c = ws.cell(row=r, column=2,
                     value=f'={eq_wt_ref}*{coe_ref}+{debt_wt_ref}*{cod_at_ref}')
    _style(wacc_c, fill_hex=LIGHT_GREEN, bold=True, h_align='right',
           number_format=FMT_PCT)
    wacc_c.border = BOT_MED
    r += 1

    ws.freeze_panes = 'A3'
    return wacc_rows


# ─────────────────────────────────────────────────────────────────────────────
# SHEET 3 – DCF MODEL
# ─────────────────────────────────────────────────────────────────────────────

