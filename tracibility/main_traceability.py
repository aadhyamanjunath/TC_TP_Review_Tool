# main_traceability.py
from __future__ import annotations
import argparse
import sys
from pathlib import Path
from typing import List, Optional

import pandas as pd
from docx import Document

import config
import io_utils as io
from traceability_logic import (
    build_bidirectional_maps,
    build_master_summary,
    check_invalid_req_ids,
    check_req_trace_coverage,
    cross_match,
)

def main():
    ap = argparse.ArgumentParser(
        description="Generate traceability report (uses defaults, CLI overrides optional)."
    )
    ap.add_argument("--req", help="Path to req .docx (optional; default hard-coded)")
    ap.add_argument("--trace_s2tc", help="Path to req→tc trace .docx (optional; default hard-coded)")
    ap.add_argument("--trace_tc2req", help="Path to tc→req trace .docx (optional; default hard-coded)")
    ap.add_argument("--out", help="Output .xlsx or directory (optional; default hard-coded)")
    args = ap.parse_args()

    placeholders_lower = {p.lower() for p in config.PLACEHOLDER_VALUES}
    req_path = Path(args.req).expanduser().resolve() if args.req else config.DEFAULT_SRS_PATH
    trace_s2tc_path = Path(args.trace_s2tc).expanduser().resolve() if args.trace_s2tc else config.DEFAULT_TRACE_TC2SREQ_PATH
    trace_tc2req_path = Path(args.trace_tc2req).expanduser().resolve() if args.trace_tc2req else config.DEFAULT_TRACE_REQ2TC_PATH
    out_path = Path(args.out).expanduser().resolve() if args.out else config.DEFAULT_OUT_XLSX
    out_path = io.ensure_out_xlsx(out_path)

    # Basic input checks
    for label, p in [("req", req_path),
                     ("Trace req→tc", trace_s2tc_path), ("Trace tc→req", trace_tc2req_path)]:
        if not p.exists():
            print(f"❌ {label} path not found: {p}")
            sys.exit(2)
        if p.suffix.lower() != ".docx":
            print(f"❌ {label} must be a .docx: {p}")
            sys.exit(2)

    CANON_REQ = config.DEFAULT_REQ_COL
    CANON_TC  = config.DEFAULT_TC_COL

    # Discover Excel inputs (avoid including the output file)
    inputs: List[Path] = io.collect_input_files(config.TC_XLSX_PATHS, config.DISCOVERY_GLOB, config.DISCOVERY_RECURSIVE, out_path)

    print("Inputs discovered:")
    for p in inputs:
        print(f"  - {p}")
    if not inputs:
        print("[ERROR] No Excel inputs found from TC_XLSX_PATHS.", file=sys.stderr)
        # Write a 'Run_Log'
        with io.open_excel_writer(out_path) as writer:
            io.add_sheet_with_col_exclusions(writer, pd.DataFrame([{
                "Status": "No Excel inputs found from TC_XLSX_PATHS"
            }]), "Run_Log")
        sys.exit(1)

    combined_frames = []
    combined_overview = []

    for src_path in inputs:
        df_expl, overview = io.process_source(
            src_path, config.DEFAULT_SHEET, config.DEFAULT_REQ_COL, config.DEFAULT_TC_COL, placeholders_lower,
            config.ENABLE_SPLIT_MULTIVALUE, config.STRIP_WRAPPING_BRACKETS, config.MULTI_VALUE_SEPARATORS,
            CANON_REQ, CANON_TC
        )
        combined_overview.append(overview)
        if overview.get("Status") == "OK" and df_expl is not None and len(df_expl) > 0:
            combined_frames.append(df_expl)

    if not combined_frames:
        print("[ERROR] No usable data from inputs to build the combined report.", file=sys.stderr)
        with io.open_excel_writer(out_path) as writer:
            io.add_sheet_with_col_exclusions(writer, pd.DataFrame(combined_overview), "Run_Log")
        sys.exit(1)

    df_long_all = pd.concat(combined_frames, ignore_index=True)
    try:
        io.write_report(
            df_long_all, combined_overview, out_path,
            CANON_REQ, CANON_TC,
            config.ENFORCE_ID_PATTERNS, config.REQ_ID_REGEX, config.TC_ID_REGEX
        )
    except Exception as e:
        print(f"[ERROR] Failed to write combined report: {e}", file=sys.stderr)
        sys.exit(1)

    # Sanity checks for later steps
    if not config.DEFAULT_SRS_PATH.exists():
        raise FileNotFoundError(f"Requirement DOCX not found: {config.DEFAULT_SRS_PATH}")

    io.ensure_out_dir(config.OUT_DIR)

    # Expand all Excel inputs for conversion to Word
    excel_files = io.expand_excel_sources(config.TC_XLSX_PATHS)
    if not excel_files:
        raise FileNotFoundError("No test case Excel files found from TC_XLSX_PATHS.")
    print("🔎 Test case Excel files to process:")
    for f in excel_files:
        print(f"   • {f}")

    # ---------------------------------------
    # EXCEL-ONLY LINK EXTRACTION (no DOCX)
    # ---------------------------------------

    # (Optional) Keep DOCX conversion just for stakeholders to read; skip parsing it
    if config.DO_BUILD_DOCX:
        excel_files = io.expand_excel_sources(config.TC_XLSX_PATHS)
        print("🔎 Test case Excel files to convert to Word (view-only):")
        for f in excel_files:
            print(f"   • {f}")

        converted_docx_paths: List[Path] = []
        combined_doc = Document() if config.COMBINED_TC_DOCX else None
        for x in excel_files:
            out_docx = config.OUT_DIR / f"{x.stem}.docx"
            print(f"\n📄 Converting: {x} → {out_docx}")
            try:
                io.excel_to_word_tables(
                    x, out_docx, sheet=config.TC_SHEET,
                    columns=config.FILTER_COLUMNS, max_rows=config.MAX_ROWS
                )
                converted_docx_paths.append(out_docx)
                print("   ✅ Done.")
            except Exception as e:
                print(f"   ❌ Failed to convert {x}: {e}", file=sys.stderr)

            if combined_doc is not None:
                try:
                    dfs = io.read_excel_sheets(x, sheet=config.TC_SHEET)
                    io.append_tables_to_doc(
                        combined_doc, dfs, title_prefix=x.name,
                        columns=config.FILTER_COLUMNS, max_rows=config.MAX_ROWS
                    )
                except Exception as e:
                    print(f"   ⚠️ Skipped adding {x.name} to combined doc: {e}", file=sys.stderr)

        if combined_doc is not None:
            try:
                combined_doc.save(config.COMBINED_TC_DOCX)   # type: ignore[arg-type]
                print(f"\n🗂️ Combined test case Word saved: {config.COMBINED_TC_DOCX.resolve()}")
            except Exception as e:
                print(f"   ❌ Failed to save combined Word: {e}", file=sys.stderr)

    # ------------------------------
    # Build links directly from Excel
    # ------------------------------
    CANON_REQ = config.DEFAULT_REQ_COL
    CANON_TC  = config.DEFAULT_TC_COL

    # 2) Parse requirements directly from Requirement DOCX (for coverage/gaps)
    print(f"\n🔍 Extracting requirements from: {config.DEFAULT_SRS_PATH}")
    req_df = io.parse_requirements(config.DEFAULT_SRS_PATH)
    req_ids = set(req_df["ID"].astype(str))
    print(f"   → Found {len(req_ids)} unique requirements in spec")

    # 3) Links from df_long_all (the exploded, normalized pairs)
    print("\n🔗 Building Requirement ↔ Test Case links from Excel rows …")
    links_df = (
        df_long_all
        .dropna(subset=[CANON_REQ, CANON_TC])
        [[CANON_REQ, CANON_TC, "Source_File"]]
        .drop_duplicates()
        .rename(columns={
            CANON_REQ: "requirement_id",
            CANON_TC:  "tc_id",
            "Source_File": "source_doc"
        })
    )

    # Rows where TC exists but Requirement is blank → "unmapped" TCs
    unmapped_tc_df = (
        df_long_all[
            df_long_all[CANON_TC].notna() & df_long_all[CANON_REQ].isna()
        ][[CANON_TC, "Source_File"]]
        .drop_duplicates()
        .rename(columns={CANON_TC: "tc_id", "Source_File": "source_doc"})
    )

    print(f"   • Links extracted: {len(links_df)} "
          f"(Reqs: {links_df['requirement_id'].nunique()}, "
          f"TCs: {links_df['tc_id'].nunique()})")
    if links_df.empty:
        print("⚠️  No links could be derived from Excel data. "
              "Verify that the canonical columns exist and contain values:\n"
              f"    - Requirement column: '{CANON_REQ}'\n"
              f"    - Test Case column : '{CANON_TC}'\n"
              "Also check that placeholders were not all filtered out.\n",
              file=sys.stderr)

        # Quick debug: show first 5 non-empty rows for either column
        debug_rows = df_long_all[
            df_long_all[CANON_REQ].notna() | df_long_all[CANON_TC].notna()
        ].head(10)
        print("🔎 Sample rows with non-empty REQ/TC:\n", debug_rows.to_string(index=False))

    # 4) Build report objects
    matrix_df = links_df.sort_values(["requirement_id", "tc_id", "source_doc"]).reset_index(drop=True)
    gaps_df = pd.DataFrame({"requirement_id": sorted(req_ids - set(matrix_df["requirement_id"].unique()))})
    orphans_df = pd.DataFrame({"requirement_id": sorted(set(matrix_df["requirement_id"].unique()) - req_ids)})

    from traceability_logic import build_bidirectional_maps
    req_to_tc, tc_to_req = build_bidirectional_maps(matrix_df)

    report = {
        "matrix": matrix_df,
        "req_to_tc": req_to_tc,
        "tc_to_req": tc_to_req,
        "gaps_in_spec": gaps_df,
        "orphans_in_tests": orphans_df,
    }
    if not unmapped_tc_df.empty:
        report["unmapped_test_cases"] = unmapped_tc_df.sort_values(["tc_id", "source_doc"]).reset_index(drop=True)

    # 5) Console summary
    print("\n=== Requirement → Test Case(s) ===")
    for _, row in report["req_to_tc"].iterrows():
        print(f"{row['requirement_id']}  →  {row['tc_ids']}  (count={row['tc_count']})")

    print("\n=== Test Case → Requirement(s) ===")
    for _, row in report["tc_to_req"].iterrows():
        print(f"{row['tc_id']}  →  {row['requirement_ids']}  (count={row['requirement_count']})")

    print("\n=== Coverage Summary ===")
    print(f"Total requirements in SPEC         : {len(req_ids)}")
    print(f"Total unique requirements in TESTS : {report['matrix']['requirement_id'].nunique()}")
    print(f"Total unique test cases            : {report['matrix']['tc_id'].nunique()}")
    print(f"Gaps (requirements in SPEC not covered) : {len(report['gaps_in_spec'])}")
    print(f"Orphans (requirements in TESTS not in SPEC): {len(report['orphans_in_tests'])}")
    if "unmapped_test_cases" in report:
        print(f"Unmapped test cases (no requirement in same row): {len(report['unmapped_test_cases'])}")

    # 6) Save workbook (excluded sheets are skipped)
    io.save_report(config.DEFAULT_OUT_XLSX, report)
    print(f"\n✅ Saved report: {config.DEFAULT_OUT_XLSX.resolve()} "
          f"(column removals applied via EXCLUDE_COLUMNS_PER_SHEET)")

    # Optional: small “consistency” workbook (Cross_Mismatch)
    try:
        trace_s2tc_df = io.parse_traceability(trace_s2tc_path)
        trace_tc2s_df = io.parse_traceability(trace_tc2req_path)

        dup_req = req_df[req_df.duplicated("ID", keep=False)]
        dup_req_rows = pd.DataFrame(
            [
                [row["ID"], "req", "DUPLICATE_ID", "Duplicate ID found in original req document"]
                for _, row in dup_req.iterrows()
            ],
            columns=["ID", "Document", "Issue_Type", "Details"]
        )
        req_coverage = check_req_trace_coverage(req_df, trace_s2tc_df)
        invalid_req_rows = check_invalid_req_ids(trace_s2tc_df, req_df)
        original_doc_issues = pd.concat([dup_req_rows, req_coverage, invalid_req_rows], ignore_index=True)

        cross = cross_match(trace_s2tc_df, trace_tc2s_df, req_df)
        master = build_master_summary(req_df, cross, req_coverage)

        with io.open_excel_writer(Path(out_path)) as writer:
            io.add_sheet_with_col_exclusions(writer, cross, "Cross_Mismatch")
            io.remove_sheets(writer, config.EXCLUDE_SHEETS)

        print("\n✔ Cross_Mismatch sheet written.")
    except Exception as e:
        print(f"\n❌ Failed to generate Cross_Mismatch: {e}", file=sys.stderr)

    print("\nInputs:")
    print(f"  req .............. {req_path}")
    print(f"  Trace req→tc .... {trace_s2tc_path}")
    print(f"  Trace tc→req .... {trace_tc2req_path}")
    print("Output:")
    print(f"  Excel report ..... {out_path}")

if __name__ == "__main__":
    main()
