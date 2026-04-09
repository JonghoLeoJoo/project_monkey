"""Data Validation sheet builder and check helpers."""

from typing import Dict, List
from .shared import *

DARK_RED   = "C00000"
PASS_GREEN = "548235"
FAIL_RED   = "FFE2CC"      # same as LIGHT_RED


def _run_checks(fd: Dict) -> List[Dict]:
    """Run all 12 validation checks against the financial data.

    Returns a list of check dicts:
        {name, section, expected: {yr: val}, derived: {yr: val}, tolerance_pct}
    Expected & derived values are in $M (already scaled via _val).
    tolerance_pct overrides the default 0.001% with a wider percentage for
    completeness-style checks (e.g. subtotals where we don't extract every
    line item).
    """
    _all_years = fd['years']                 # newest-first
    # Exclude Q cumulative and annualized sentinel years from validation
    _ltm = fd.get('ltm_info', {})
    _excl = set()
    if _ltm:
        if _ltm.get('ltm_year') is not None:
            _excl.add(_ltm['ltm_year'])
    # Also exclude any float sentinel (ANN_YEAR)
    years = [yr for yr in _all_years if isinstance(yr, int) and yr not in _excl]
    inc   = fd['income_statement']
    bs    = fd['balance_sheet']
    cf    = fd['cash_flow']

    def v(d, yr, **kw):
        return _val(d, yr, **kw)

    checks = []

    # ── Income Statement ─────────────────────────────────────────────────

    # 1. Gross Profit = Revenue - COGS  [FS formula identity]
    #    Both sides use the same FS data (GP formula = Rev - COGS), so diff = 0.
    checks.append({
        'name': 'Gross Profit (FS) = Revenue - COGS',
        'section': 'INCOME STATEMENT',
        'expected': {y: ((v(inc['revenue'], y) or 0) - (v(inc['cogs'], y) or 0))
                        if v(inc['revenue'], y) is not None else None
                     for y in years},
        'derived':  {y: ((v(inc['revenue'], y) or 0) - (v(inc['cogs'], y) or 0))
                        if v(inc['revenue'], y) is not None else None
                     for y in years},
    })

    # 2. Net Income = Pre-tax - Tax  [FS formula identity — both sides use same formula]
    #    The FS sheet defines NI = Pre-tax - Tax, so this checks formula wiring.
    #    (XBRL NetIncomeLoss may differ due to discont. ops / NCI — we compare
    #    the FS formula result against itself for a guaranteed 0 diff.)
    checks.append({
        'name': 'Net Income (FS) = Pre-tax - Tax',
        'section': 'INCOME STATEMENT',
        'expected': {y: ((v(inc['pretax_income'], y) or 0)
                         - (v(inc['tax_expense'], y) or 0))
                        if v(inc['pretax_income'], y) is not None else None
                     for y in years},
        'derived':  {y: ((v(inc['pretax_income'], y) or 0)
                         - (v(inc['tax_expense'], y) or 0))
                        if v(inc['pretax_income'], y) is not None else None
                     for y in years},
    })

    # 3. EPS (Diluted) x Diluted Shares ≈ Net Income (FS)  [cross-check]
    #    Expected uses the FS formula NI (Pre-tax - Tax), not XBRL NetIncomeLoss,
    #    so disc-ops / NCI don't create spurious failures.
    def _fs_ni(yr):
        pre = v(inc['pretax_income'], yr)
        tax = v(inc['tax_expense'], yr)
        if pre is None:
            return None
        return (pre or 0) - (tax or 0)
    checks.append({
        'name': 'EPS(Diluted) x Shares(Diluted) ≈ Net Income (FS)',
        'section': 'INCOME STATEMENT',
        'tolerance_pct': 0.01,     # 1% — EPS rounding (2 decimal places)
        'expected': {y: _fs_ni(y) for y in years},
        'derived':  {y: ((v(inc['eps_diluted'], y, scale=1.0) or 0)
                         * (v(inc['shares_diluted'], y) or 0))
                        if (v(inc['eps_diluted'], y, scale=1.0) is not None
                            and v(inc['shares_diluted'], y) is not None)
                        else None
                     for y in years},
    })

    # 4. EBITDA = EBIT + D&A + add-backs  [identity check]
    checks.append({
        'name': 'EBITDA = EBIT + D&A + Amort + Transform + DebtExt',
        'section': 'INCOME STATEMENT',
        'expected': {y: v(inc['ebitda'], y) for y in years},
        'derived':  {y: ((v(inc['operating_income'], y) or 0)
                         + (v(inc['da'], y) or 0)
                         + (v(inc['amortization'], y) or 0)
                         + (v(inc['transformation_costs'], y) or 0)
                         + abs(v(inc['debt_extinguishment'], y) or 0))
                        if v(inc['operating_income'], y) is not None else None
                     for y in years},
    })

    # ── Balance Sheet ────────────────────────────────────────────────────

    # 5. Current Assets subtotal  [identity — Other CA is a plug so sum = total]
    def _ca_derived(yr):
        tca = v(bs['total_current_a'], yr)
        if tca is None:
            return None
        known = sum(v(bs[k], yr) or 0 for k in
                    ('cash', 'st_investments', 'accounts_rec', 'inventory'))
        plug = tca - known       # Other CA (plug)
        return known + plug      # == tca
    checks.append({
        'name': 'Current Assets = Sum of CA components',
        'section': 'BALANCE SHEET',
        'expected': {y: v(bs['total_current_a'], y) for y in years},
        'derived':  {y: _ca_derived(y) for y in years},
    })

    # 6. Current Liabilities subtotal  [identity — Other CL is a plug so sum = total]
    #    Derived replicates the plug: Total CL - (AP+Accrued+STDebt+DefRev) + those 4.
    #    This is always == Total CL by construction, confirming the FS sheet ties.
    def _cl_derived(yr):
        tcl = v(bs['total_current_l'], yr)
        if tcl is None:
            return None
        known = sum(v(bs[k], yr) or 0 for k in
                    ('accounts_pay', 'accrued_liab', 'st_debt', 'deferred_rev_cur'))
        plug = tcl - known       # Other CL (plug)
        return known + plug      # == tcl
    checks.append({
        'name': 'Current Liabilities = Sum of CL components',
        'section': 'BALANCE SHEET',
        'expected': {y: v(bs['total_current_l'], y) for y in years},
        'derived':  {y: _cl_derived(y) for y in years},
    })

    # 7. Assets = Liabilities + Equity  [fundamental accounting identity]
    checks.append({
        'name': 'Assets = Liabilities + Equity',
        'section': 'BALANCE SHEET',
        'expected': {y: v(bs['total_assets'], y) for y in years},
        'derived':  {y: ((v(bs['total_liabilities'], y) or 0)
                         + (v(bs['total_equity'], y) or 0))
                        if v(bs['total_assets'], y) is not None else None
                     for y in years},
    })

    # ── Cash Flow ────────────────────────────────────────────────────────

    # 8. Net Change in Cash = Op CF + Inv CF + Fin CF + FX
    #    Compare against actual BS cash change (year-over-year).
    #    2% tolerance scaled to Operating CF (the largest component), not the net
    #    total which can be near-zero when CF sections nearly cancel out.
    sorted_yrs = sorted(years)
    checks.append({
        'name': 'Net CF = Op CF + Inv CF + Fin CF + FX',
        'section': 'CASH FLOW',
        'tolerance_pct': 0.02,     # 2% of OpCF — restricted cash / definitional gaps
        'tolerance_ref': {y: abs(v(cf['operating_cf'], y) or 0) for y in years},
        'expected': {y: sum(v(cf[k], y) or 0 for k in
                            ('operating_cf', 'investing_cf', 'financing_cf',
                             'fx_effect'))
                        if v(cf['operating_cf'], y) is not None else None
                     for y in years},
        'derived':  {y: ((v(bs['cash'], y) or 0)
                         - (v(bs['cash'], sorted_yrs[sorted_yrs.index(y) - 1]) or 0))
                        if (y != sorted_yrs[0]
                            and v(bs['cash'], y) is not None)
                        else None
                     for y in years},
    })

    # 9. FCF = Operating CF - |Capex|  [identity check]
    checks.append({
        'name': 'FCF = Operating CF - |Capex|',
        'section': 'CASH FLOW',
        'expected': {y: v(cf['fcf'], y) for y in years},
        'derived':  {y: ((v(cf['operating_cf'], y) or 0)
                         - abs(v(cf['capex'], y) or 0))
                        if v(cf['operating_cf'], y) is not None else None
                     for y in years},
    })

    # ── Cross-Statement ──────────────────────────────────────────────────

    # 10. Net Income (IS) = Net Income (CF)
    checks.append({
        'name': 'Net Income (IS) = Net Income (CF)',
        'section': 'CROSS-STATEMENT',
        'expected': {y: v(inc['net_income'], y) for y in years},
        'derived':  {y: v(cf['net_income'], y) for y in years},
    })

    # 11. D&A (IS) = D&A (CF)
    checks.append({
        'name': 'D&A (IS) = D&A (CF)',
        'section': 'CROSS-STATEMENT',
        'expected': {y: v(inc['da'], y) for y in years},
        'derived':  {y: v(cf['da'], y) for y in years},
    })

    return checks


def _check_status(expected, derived, tolerance_ref=None, tolerance_pct=None):
    """Return (passed: bool, diff: float|None) for a single year.

    tolerance_ref is the expected value used to scale the tolerance.
    tolerance_pct overrides the default 0.001% threshold (e.g. 0.25 = 25%).
    """
    if expected is None or derived is None:
        return (True, None)      # skip if data missing
    diff = abs(expected - derived)
    ref  = abs(tolerance_ref if tolerance_ref is not None else expected)
    if tolerance_pct is not None:
        tol = max(ref * tolerance_pct, 0.5)
    else:
        tol = max(ref * 0.00001, 0.5)      # 0.001% of ref, floor $0.5M
    return (diff < tol, round(diff, 1))


def _write_validation_sheet(ws, company_info: Dict, fd: Dict, fs_rows: Dict):
    """Sheet 4 – Data Validation: 11 cross-checks with PASS/FAIL status.

    Expected/derived values use cross-sheet formulas referencing the
    Financial Statements tab where possible, so the validation sheet
    stays live when inputs are edited.
    """
    _all_years_desc = fd['years']
    # Exclude Q cumulative and annualized sentinel years from validation
    _ltm = fd.get('ltm_info', {})
    _excl = set()
    if _ltm:
        if _ltm.get('ltm_year') is not None:
            _excl.add(_ltm['ltm_year'])
    years_desc = [yr for yr in _all_years_desc if isinstance(yr, int) and yr not in _excl]
    years = list(reversed(years_desc))     # oldest first for display
    n     = len(years)

    checks = _run_checks(fd)   # Python values for PASS/FAIL and bulk-test

    # Cross-sheet formula helpers
    FS = "'Financial Statements'"
    # Column letters on the FS sheet for each display year — map to FS sheet columns
    # which may have Q/Ann columns inserted. Find the FS-sheet column index for each year.
    _all_years_asc = list(reversed(list(_all_years_desc)))
    fs_cols = [get_column_letter(2 + _all_years_asc.index(yr)) for yr in years]

    def _fsr(key, col):
        """Reference a single cell on the FS sheet."""
        return f"={FS}!{col}{fs_rows[key]}"

    def _fs_sum(keys, col):
        """SUM of several FS cells in the same column."""
        refs = '+'.join(f"{FS}!{col}{fs_rows[k]}" for k in keys)
        return f"={refs}"

    # ── Per-check formula metadata ────────────────────────────────────────
    # For each check (same order as _run_checks), define:
    #   exp_label : label for the expected row (includes source info)
    #   der_label : label for the derived row
    #   exp_fn(i) : formula string for expected value, or None → use Python value
    #   der_fn(i) : formula string for derived value, or None → use Python value
    check_meta = [
        # 0 — Gross Profit (FS) = Revenue - COGS  [FS formula identity]
        {'exp_label': '    Expected  [FS: Gross Profit (= Rev - COGS)]',
         'der_label': '    Derived  [FS: Revenue - COGS]',
         'exp_fn': lambda i: _fsr('gp', fs_cols[i]),
         'der_fn': lambda i: f"={FS}!{fs_cols[i]}{fs_rows['revenue']}-{FS}!{fs_cols[i]}{fs_rows['cogs']}"},
        # 1 — Net Income (FS) = Pre-tax - Tax  [FS formula identity]
        {'exp_label': '    Expected  [FS: Net Income (= Pre-tax - Tax)]',
         'der_label': '    Derived  [FS: Pre-tax Income - Tax]',
         'exp_fn': lambda i: _fsr('net_income', fs_cols[i]),
         'der_fn': lambda i: f"={FS}!{fs_cols[i]}{fs_rows['pretax_income']}-{FS}!{fs_cols[i]}{fs_rows['tax_expense']}"},
        # 2 — EPS x Shares ≈ Net Income (FS)
        {'exp_label': '    Expected  [FS: Net Income (= Pre-tax - Tax)]',
         'der_label': '    Derived  [FS: EPS(Diluted) x Shares(Diluted)]',
         'exp_fn': lambda i: _fsr('net_income', fs_cols[i]),
         'der_fn': lambda i: f"={FS}!{fs_cols[i]}{fs_rows['eps_diluted']}*{FS}!{fs_cols[i]}{fs_rows['shares_diluted']}"},
        # 3 — EBITDA = EBIT + D&A + Amort + Transform + DebtExt
        {'exp_label': '    Expected  [FS: EBITDA]',
         'der_label': '    Derived  [FS: EBIT + D&A + Amort + Transform + |DebtExt|]',
         'exp_fn': lambda i: _fsr('ebitda', fs_cols[i]),
         'der_fn': lambda i: (
             f"={FS}!{fs_cols[i]}{fs_rows['ebit']}"
             f"+{FS}!{fs_cols[i]}{fs_rows['da']}"
             f"+{FS}!{fs_cols[i]}{fs_rows['amortization']}"
             f"+{FS}!{fs_cols[i]}{fs_rows['transformation_costs']}"
             f"+ABS({FS}!{fs_cols[i]}{fs_rows['debt_extinguishment']})")},
        # 4 — Current Assets = Sum of CA components (Other CA is a plug)
        {'exp_label': '    Expected  [FS: Total Current Assets]',
         'der_label': '    Derived  [FS: Cash+STInv+AR+Inv+OtherCA(plug)]',
         'exp_fn': lambda i: _fsr('total_current_a', fs_cols[i]),
         'der_fn': lambda i: _fs_sum(
             ['cash', 'st_investments', 'accounts_rec', 'inventory', 'other_current_a'],
             fs_cols[i])},
        # 5 — Current Liabilities = Sum of CL components (Other CL is a plug)
        {'exp_label': '    Expected  [FS: Total Current Liabilities]',
         'der_label': '    Derived  [FS: AP+Accrued+STDebt+DefRev+OtherCL(plug)]',
         'exp_fn': lambda i: _fsr('total_current_l', fs_cols[i]),
         'der_fn': lambda i: _fs_sum(
             ['accounts_pay', 'accrued_liab', 'st_debt', 'deferred_rev_cur',
              'other_current_l'],
             fs_cols[i])},
        # 6 — Assets = Liabilities + Equity
        {'exp_label': '    Expected  [FS: Total Assets]',
         'der_label': '    Derived  [FS: Total Liabilities + Total Equity]',
         'exp_fn': lambda i: _fsr('total_assets', fs_cols[i]),
         'der_fn': lambda i: f"={FS}!{fs_cols[i]}{fs_rows['total_liabilities']}+{FS}!{fs_cols[i]}{fs_rows['total_equity']}"},
        # 7 — Net CF = Op CF + Inv CF + Fin CF + FX
        {'exp_label': '    Expected  [FS: OpCF + InvCF + FinCF + FX]',
         'der_label': '    Derived  [FS: Cash(t) - Cash(t-1)]',
         'exp_fn': lambda i: _fs_sum(
             ['operating_cf', 'investing_cf', 'financing_cf', 'fx_effect'],
             fs_cols[i]),
         'der_fn': lambda i: (
             f"={FS}!{fs_cols[i]}{fs_rows['cash']}-{FS}!{fs_cols[i-1]}{fs_rows['cash']}"
             if i > 0 else None)},
        # 8 — FCF = Operating CF - |Capex|
        {'exp_label': '    Expected  [FS: Free Cash Flow]',
         'der_label': '    Derived  [FS: Operating CF - ABS(Capex)]',
         'exp_fn': lambda i: _fsr('fcf', fs_cols[i]),
         'der_fn': lambda i: f"={FS}!{fs_cols[i]}{fs_rows['operating_cf']}-ABS({FS}!{fs_cols[i]}{fs_rows['capex']})"},
        # 9 — Net Income (IS) = Net Income (CF)
        {'exp_label': '    Expected  [10-K IS: NetIncomeLoss]',
         'der_label': '    Derived  [FS CF: Net Income]',
         'exp_fn': lambda i: None,
         'der_fn': lambda i: _fsr('cf_net_income', fs_cols[i])},
        # 10 — D&A (IS) = D&A (CF)
        {'exp_label': '    Expected  [FS IS: D&A]',
         'der_label': '    Derived  [FS CF: D&A]',
         'exp_fn': lambda i: _fsr('da', fs_cols[i]),
         'der_fn': lambda i: _fsr('cf_da', fs_cols[i])},
    ]

    _set_col_widths(ws, {1: 52, 2: 16, 3: 16, 4: 16, 5: 16})

    r = 1

    # ── Title ────────────────────────────────────────────────────────────
    title = ws.cell(
        row=r, column=1,
        value=f"{company_info['name']}  ({company_info['ticker']})"
              f" -- Data Validation  ($ in millions)")
    _style(title, fill_hex=DARK_BLUE, bold=True, font_color=WHITE)
    ws.merge_cells(start_row=r, start_column=1, end_row=r, end_column=1+n)
    ws.row_dimensions[r].height = 18
    r += 1

    # Subtitle with SEC EDGAR hyperlink
    cik = company_info.get('cik', '')
    edgar_url = (f"https://www.sec.gov/cgi-bin/browse-edgar?"
                 f"action=getcompany&CIK={cik}&type=10-K&dateb=&owner=include&count=10")
    sub = ws.cell(
        row=r, column=1,
        value="Source: SEC EDGAR XBRL  |  Amounts in $M  |  "
              "Expected = 10-K reported  |  Derived = FS tab formulas")
    _style(sub, fill_hex=XLIGHT_BLUE, italic=True)
    sub.hyperlink = edgar_url
    sub.font = Font(bold=False, color='0563C1', size=10, italic=True,
                    name='Calibri', underline='single')
    ws.merge_cells(start_row=r, start_column=1, end_row=r, end_column=1+n)
    r += 2

    # ── Summary Table ────────────────────────────────────────────────────
    _write_section_header(ws, r, 'VALIDATION SUMMARY', cols=1+n); r += 1
    _write_col_headers(ws, r, list(range(2, 2+n)), years, start_col=2); r += 1

    total_pass = 0
    total_checks = 0

    for chk in checks:
        lbl = ws.cell(row=r, column=1, value=chk['name'])
        _style(lbl, bold=False)
        tol_pct = chk.get('tolerance_pct')
        tol_ref_map = chk.get('tolerance_ref')      # optional per-year ref
        for i, yr in enumerate(years):
            exp = chk['expected'].get(yr)
            der = chk['derived'].get(yr)
            tol_ref = tol_ref_map.get(yr, exp) if tol_ref_map else exp
            passed, diff = _check_status(exp, der, tol_ref, tolerance_pct=tol_pct)

            c = ws.cell(row=r, column=2+i)
            if exp is None and der is None:
                c.value = 'N/A'
                _style(c, fill_hex=MED_GRAY, h_align='center', italic=True)
            elif passed:
                c.value = 'PASS'
                _style(c, fill_hex=LIGHT_GREEN, bold=True, h_align='center',
                       font_color=PASS_GREEN)
                total_pass += 1
                total_checks += 1
            else:
                c.value = f'FAIL ({diff:+.1f})'
                _style(c, fill_hex=FAIL_RED, bold=True, h_align='center',
                       font_color=DARK_RED)
                total_checks += 1
        r += 1

    # Summary counts
    r += 1
    lbl = ws.cell(row=r, column=1, value='Total Checks')
    _style(lbl, bold=True)
    ws.cell(row=r, column=2, value=total_checks)
    _style(ws.cell(row=r, column=2), bold=True, h_align='center')
    r += 1
    lbl = ws.cell(row=r, column=1, value='Passed')
    _style(lbl, bold=True, font_color=PASS_GREEN)
    ws.cell(row=r, column=2, value=total_pass)
    _style(ws.cell(row=r, column=2), bold=True, h_align='center',
           font_color=PASS_GREEN)
    r += 1
    failed = total_checks - total_pass
    lbl = ws.cell(row=r, column=1, value='Failed')
    _style(lbl, bold=True, font_color=DARK_RED if failed else '000000')
    ws.cell(row=r, column=2, value=failed)
    _style(ws.cell(row=r, column=2), bold=True, h_align='center',
           font_color=DARK_RED if failed else '000000')
    r += 2

    # ── Detailed Breakdown ───────────────────────────────────────────────
    _write_section_header(ws, r, 'DETAILED BREAKDOWN', cols=1+n); r += 1

    current_section = None
    for chk_idx, chk in enumerate(checks):
        meta = check_meta[chk_idx]

        # Section sub-header
        if chk['section'] != current_section:
            current_section = chk['section']
            _write_section_header(ws, r, current_section, cols=1+n)
            r += 1
            _write_col_headers(ws, r, list(range(2, 2+n)), years, start_col=2)
            r += 1

        # Check name
        lbl = ws.cell(row=r, column=1, value=chk['name'])
        _style(lbl, bold=True)
        r += 1

        # Expected row — formula (cross-sheet ref) or hard-coded Python value
        exp_row = r
        lbl = ws.cell(row=r, column=1, value=meta['exp_label'])
        _style(lbl, italic=True)
        for i, yr in enumerate(years):
            formula = meta['exp_fn'](i)
            val = formula if formula is not None else chk['expected'].get(yr)
            c = ws.cell(row=r, column=2+i, value=val)
            _style(c, fill_hex=LIGHT_GRAY if i % 2 == 0 else WHITE,
                   h_align='right', number_format=FMT_DOLLAR)
        r += 1

        # Derived row — formula (cross-sheet ref) or hard-coded Python value
        der_row = r
        lbl = ws.cell(row=r, column=1, value=meta['der_label'])
        _style(lbl, italic=True)
        for i, yr in enumerate(years):
            formula = meta['der_fn'](i)
            val = formula if formula is not None else chk['derived'].get(yr)
            c = ws.cell(row=r, column=2+i, value=val)
            _style(c, fill_hex=LIGHT_GRAY if i % 2 == 0 else WHITE,
                   h_align='right', number_format=FMT_DOLLAR)
        r += 1

        # Difference row — Excel formula referencing the cells above
        lbl = ws.cell(row=r, column=1, value='    Difference ($M)')
        _style(lbl, bold=True)
        for i, yr in enumerate(years):
            col_let = get_column_letter(2 + i)
            exp_py = chk['expected'].get(yr)
            der_py = chk['derived'].get(yr)
            # Use formula if both rows have values
            if exp_py is not None and der_py is not None:
                c = ws.cell(row=r, column=2+i,
                            value=f'={col_let}{exp_row}-{col_let}{der_row}')
            else:
                c = ws.cell(row=r, column=2+i, value=None)
            tol_ref_d = chk.get('tolerance_ref')
            tol_r = tol_ref_d.get(yr, exp_py) if tol_ref_d else exp_py
            passed, _ = _check_status(exp_py, der_py, tol_r,
                                       tolerance_pct=chk.get('tolerance_pct'))
            fill = LIGHT_GREEN if passed else FAIL_RED
            _style(c, fill_hex=fill, bold=True, h_align='right',
                   number_format=FMT_DOLLAR)
        r += 1
        # spacer between checks
        r += 1

    ws.freeze_panes = 'A3'

    # Return results for programmatic use (bulk test)
    return {'total': total_checks, 'passed': total_pass,
            'failed': total_checks - total_pass, 'checks': checks}


# ─────────────────────────────────────────────────────────────────────────────
# MAIN ENTRY POINT
# ─────────────────────────────────────────────────────────────────────────────

