#!/usr/bin/env python3
# -*- coding: utf-8 -*-

"""DO-178 Test Case Review Tool - CLI entry point

Runs Environment/Equipment cell checks per row and writes an Excel/CSV report.

Usage examples:
  python main.py -t testcases.xlsx -r ref.docx -o review_report.xlsx
  python -m yourpackage.main -t testcases.xlsx -r ref.docx
"""

import argparse
import sys

# Support both package execution (relative import) and direct script execution.
try:
    from .logic import (
        load_docx_text,
        extract_section_numbers_from_docx,
        load_testcases,
        resolve_env_column,
        load_equipment_keywords,
        review_env_cells,
        save_report,
    )
except ImportError:  # pragma: no cover
    from logic import (
        load_docx_text,
        extract_section_numbers_from_docx,
        load_testcases,
        resolve_env_column,
        load_equipment_keywords,
        review_env_cells,
        save_report,
    )


def main() -> None:
    ap = argparse.ArgumentParser(
        description="DO-178 Env/Equipment review tool (cell-level checks per row)"
    )
    ap.add_argument("--testcases", "-t", required=True,
                    help="Path to testcase file (.xlsx/.xls/.csv)")
    ap.add_argument("--sheet", "-s", default=None, help="Sheet name (Excel only)")
    ap.add_argument("--refdoc", "-r", required=True, help="Reference Word doc (.docx)")
    ap.add_argument("--out", "-o", default="review_report.xlsx",
                    help="Output report path (.xlsx or .csv)")
    ap.add_argument("--col-env", default="Test Environment/Equipment",
                    help="Environment or Equipment column name in the sheet")
    ap.add_argument("--header-row", type=int, default=None,
                    help="0-based header row (Excel). Auto-detect if omitted.")
    ap.add_argument("--equipment-list", default=None,
                    help="Optional equipment keywords file (.txt/.csv)")

    args = ap.parse_args()

    # Load reference document data
    doc_text = load_docx_text(args.refdoc)
    section_numbers = extract_section_numbers_from_docx(args.refdoc)

    # Load test cases
    df = load_testcases(args.testcases, args.sheet, args.header_row, args.col_env)
    env_col = resolve_env_column(df, args.col_env)

    # Load equipment keywords
    equip_keywords = load_equipment_keywords(args.equipment_list)

    # Review
    reviewed, summary, df_empty = review_env_cells(
        df=df,
        env_col=env_col,
        doc_text=doc_text,
        section_numbers=section_numbers,
        equip_keywords=equip_keywords,
        header_row=args.header_row,
    )

    out_path = save_report(reviewed, summary, df_empty, args.out)

    # Console summary
    print("=== DO-178 Environment/Equipment Review Summary ===")
    print(f"Environment column used: {env_col}")
    for k, v in summary.items():
        print(f"{k}: {v}")
    print(f"Report written to: {out_path}")


if __name__ == "__main__":
    try:
        main()
    except Exception as e:
        print(f"[ERROR] {e}", file=sys.stderr)
        sys.exit(1)
