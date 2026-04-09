"""Step 6 – Excel ワークブックのビルドと JSON 保存."""

import json
import os
from typing import Dict, List, Optional

from excel import create_excel


def _safe_name(name: str) -> str:
    return (name
            .replace(' ', '_')
            .replace('/', '-')
            .replace('\\', '-')
            .replace('.', '')
            .replace(',', ''))


def _stringify_keys(obj):
    """JSON は文字列キーのみを受け付けるため、int キーを変換する."""
    if isinstance(obj, dict):
        return {str(k): _stringify_keys(v) for k, v in obj.items()}
    return obj


def run(company_info: Dict, financial_data: Dict, years: List[int]) -> Optional[Dict]:
    """Excel ファイルを生成し、10-K データを JSON で保存する.

    Returns:
        validation_results dict, or None
    """
    print()
    print("[6/6] Building Excel financial model...")

    safe_name  = _safe_name(company_info['name'])
    models_dir = os.path.join(os.path.dirname(os.path.abspath(__file__)), '..', 'models')
    output_dir = os.path.normpath(models_dir)
    os.makedirs(output_dir, exist_ok=True)
    output_file = os.path.join(output_dir, f"{safe_name}_Financial_Model.xlsx")

    tenk_dir = os.path.join(output_dir, '10k')
    os.makedirs(tenk_dir, exist_ok=True)

    tenk_file = os.path.join(tenk_dir, f"{safe_name}_10k.json")
    with open(tenk_file, 'w', encoding='utf-8') as f:
        json.dump({
            'company':        company_info,
            'fiscal_years':   years,
            'financial_data': _stringify_keys(financial_data),
        }, f, indent=2)
    print(f"  [OK] 10-K data saved to: models/10k/{safe_name}_10k.json")

    validation_results = create_excel(company_info, financial_data, output_file)

    abs_path = os.path.abspath(output_file)
    print(f"\n{'=' * 60}\n  Done!\n  File saved to: {abs_path}\n{'=' * 60}\n")
    print(
        "  Excel workbook contents:\n"
        "  - Sheet 'Financial Statements': 3-statement model (historical)\n"
        "  - Sheet 'WACC'               : Weighted avg cost of capital\n"
        "  - Sheet 'DCF Model'           : 5-year DCF valuation\n"
        "  - Sheet 'Data Validation'     : 12 cross-checks (PASS/FAIL)\n"
    )
    if validation_results:
        p = validation_results['passed']
        t = validation_results['total']
        f = validation_results['failed']
        print(f"  Validation: {p}/{t} checks passed", end='')
        if f > 0:
            print(f"  ({f} failed - see Data Validation sheet)")
        else:
            print("  (all passed)")
        print(
            "\n  Tip: Yellow cells in the WACC and DCF sheets are editable inputs.\n"
            "       Change WACC assumptions and growth rates to run scenarios.\n"
        )

    return validation_results
