#!/usr/bin/env python3
"""main.py

CLI entrypoint for the Stack/Overflow checkpoint review tool.

Example:
    python main.py --srs path/to/SRS.docx --trace Traceability.xlsx --tc TestCases.xlsx --out Report.xlsx
"""

from __future__ import annotations

import argparse
from pathlib import Path

from logic import (
    DEFAULT_KEYWORDS,
    DEFAULT_OUT,
    DEFAULT_REQ_ID_REGEX,
    DEFAULT_SRS,
    DEFAULT_TC,
    DEFAULT_TRACE,
    run,
)


def build_parser() -> argparse.ArgumentParser:
    parser = argparse.ArgumentParser(
        description="Automated reviewer: SRS -> Traceability -> TC scenario review (stack checkpoint)",
        formatter_class=argparse.ArgumentDefaultsHelpFormatter,
    )

    parser.add_argument("--srs", default=str(DEFAULT_SRS), help="SRS document path (.docx)")
    parser.add_argument("--trace", default=str(DEFAULT_TRACE), help="Traceability matrix path (.xlsx)")
    parser.add_argument("--tc", default=str(DEFAULT_TC), help="Test cases Excel path (.xlsx)")
    parser.add_argument("--out", default=str(DEFAULT_OUT), help="Output Excel path (.xlsx)")

    parser.add_argument("--keywords", default=DEFAULT_KEYWORDS, help="Semicolon-separated keywords")
    parser.add_argument("--req_id_regex", default=DEFAULT_REQ_ID_REGEX, help="Regex pattern for requirement IDs")
    parser.add_argument("--trace_sheet", default=None, help="Traceability sheet name/index (optional)")
    parser.add_argument("--tc_sheet", default="Test Cases", help="TC sheet name/index (default: 'Test Cases')")

    return parser


def main(argv=None) -> int:
    parser = build_parser()
    args = parser.parse_args(argv)

    run(
        srs=Path(args.srs),
        trace=Path(args.trace),
        tc=Path(args.tc),
        out=Path(args.out),
        keywords=args.keywords,
        req_id_regex=args.req_id_regex,
        trace_sheet=args.trace_sheet,
        tc_sheet=args.tc_sheet,
    )

    return 0


if __name__ == "__main__":
    raise SystemExit(main())

