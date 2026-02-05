# =========================================================
# CONFIGURATION FILE – G6 TRACEABILITY
# =========================================================
# config.py
from __future__ import annotations
from pathlib import Path
from typing import List, Optional, Dict, Set

# ---------------------------
# DEFAULT INPUT/OUTPUT PATHS
# ---------------------------
DEFAULT_SRS_PATH = Path(r"C:\Users\nkr\TC_TP_Tracability\SCU_Full_SRS.docx")
DEFAULT_TRACE_TC2SREQ_PATH = Path(r"C:\Users\nkr\TC_TP_Tracability\TC_Req.docx")
DEFAULT_TRACE_REQ2TC_PATH = Path(r"C:\Users\nkr\TC_TP_Tracability\Req_TC.docx")

# One or more Excel sources (files or folders)
TC_XLSX_PATHS: List[str] = [r"C:\Users\nkr\TC_TP_Tracability\TC_TP"]

# Output directory (created if missing)
OUT_DIR = Path(r"C:\Users\nkr\TC_TP_Tracability\out")

# Optional combined Word with all Excel tables (set to None to skip)
COMBINED_TC_DOCX: Optional[Path] = OUT_DIR / "All_TestCases_Combined.docx"

# Final Excel report
DEFAULT_OUT_XLSX = Path(r"C:\Users\nkr\TC_TP_Tracability\Traceability_Result.xlsx")

# Discovery options for folders:
DISCOVERY_GLOB = "*.xls*"     # includes .xls, .xlsx
DISCOVERY_RECURSIVE = True

# Optional: restrict sheets, columns, rows during conversion (keep None for all)
TC_SHEET: Optional[str | int] = None                 # e.g., "Sheet1" or 0 (None = all sheets)
FILTER_COLUMNS: Optional[List[str]] = None           # e.g., ["Requirement_ID", "TC_ID", "Scenario_No"]
MAX_ROWS: Optional[int] = None                       # e.g., 1000

def _ensure_out_path(p: Path) -> Path:
    """Ensure output is an .xlsx and the directory exists."""
    p = p if p.suffix.lower() == ".xlsx" else p.with_suffix(".xlsx")
    if p.parent and not p.parent.exists():
        p.parent.mkdir(parents=True, exist_ok=True)
    return p

# =========================
# DEFAULTS & TOGGLES
# =========================

# Set to False to bypass DOCX completely and derive links directly from Excel rows.
DO_BUILD_DOCX = False

# Sheet/columns (case-insensitive resolution supported)
DEFAULT_SHEET = 0                     # index (0) or sheet name string (e.g., "Sheet1")
DEFAULT_REQ_COL = "Requirement_ID"
DEFAULT_TC_COL = "TC_ID"

# Placeholders -> treated as NA (case-insensitive)
PLACEHOLDER_VALUES = ["unchanged", "n/a", "na", "none", "-"]

# Multi-valued cells
ENABLE_SPLIT_MULTIVALUE = True
MULTI_VALUE_SEPARATORS = r"[;,|]"    # regex; e.g., "TC-1, TC-2 | TC-3"
STRIP_WRAPPING_BRACKETS = True       # remove surrounding [] or ()

# --- Project-specific prefixes (tune here) ---
PROJECT_PREFIXES = ["MRJ_SCU_STC"]  # add more if needed

# ---- ID pattern strings (used for value-based detection/validation)
_prefix_group = "|".join(PROJECT_PREFIXES)
REQ_SRS_STRICT = rf"^(?:{_prefix_group})_SRS_\d{{1,6}}$"
REQ_SYS_STRICT = rf"^(?:{_prefix_group})_SYS_\d{{1,6}}$"
REQ_SRS_SIMPLE = r"^SRS[-_ ]?\d{1,6}$"
REQ_SYS_SIMPLE = r"^SYS[-_ ]?\d{1,6}$"
REQ_GENERIC_FALLBACK = r"^[A-Z0-9]{2,}(?:_[A-Z0-9]{2,})+_\d{1,6}$"
REQ_ID_REGEX = rf"(?:{REQ_SRS_STRICT}|{REQ_SYS_STRICT}|{REQ_SRS_SIMPLE}|{REQ_SYS_SIMPLE}|{REQ_GENERIC_FALLBACK})"

TC_PREFIXED = r"^[Tt]\s*[Cc]\s*[-_ ]?\d{1,6}$"
STC_PREFIXED = r"^[Ss]\s*[Tt]\s*[Cc]\s*[-_ ]?\d{1,6}$"
TC_TEST_WORD = r"^[Tt]est(?:\s*[Cc]ase)?[-_ ]?\d{1,6}$"
TC_ID_REGEX = rf"(?:{TC_PREFIXED}|{STC_PREFIXED}|{TC_TEST_WORD})"

# Optional ID format validation (regex)
ENFORCE_ID_PATTERNS = False  # set True to enforce the above patterns

# Candidate regexes for detection
REQ_ID_CANDIDATE_REGEXES = [REQ_SRS_STRICT, REQ_SYS_STRICT, REQ_SRS_SIMPLE, REQ_SYS_SIMPLE, REQ_GENERIC_FALLBACK]
TC_ID_CANDIDATE_REGEXES = [TC_PREFIXED, STC_PREFIXED, TC_TEST_WORD]

# Header word hints
REQ_HEADER_HINTS = [
    "req", "require", "requirement", "hlr", "srs", "sys",
    "req id", "requirement id", "hlr id", "srs id", "sys id",
    "reqid", "requirementid", "req_id", "requirement_id",
    "requirement no", "req no", "requirement number"
]
TC_HEADER_HINTS = [
    "tc", "test", "case", "testcase", "tcs", "tcid", "test id",
    "tc id", "testcase id", "test case id", "tc_no",
    "tc number", "tc no", "testcaseid", "test_case_id"
]
ID_SYNONYMS = ["id", "no", "number", "identifier"]
REQ_NEGATIVE_HINTS = ["tc", "test", "testcase", "tp", "step"]
TC_NEGATIVE_HINTS = ["req", "require", "requirement", "hlr", "srs", "sys"]

# ---------------- SHEETS TO EXCLUDE & REMOVAL ----------------
EXCLUDE_SHEETS: Set[str] = {
    "Summary",
    "Sources_Overview",
    "Req_to_TC_Map",
    "Req_without_TC",
    "matrix",
    "orphans_in_tests",
    "unmapped_test_cases",
    "Original_Document_Issues",
    "Master_Traceability_Summary",
    "Traceability_Issues",
}

# ---------------- COLUMNS TO EXCLUDE PER SHEET ----------------
EXCLUDE_COLUMNS_PER_SHEET: Dict[str, set[str]] = {
    "Exploded_With_Source": {"Source_File", "Source_Sheet"},
    "Req_to_TC_Map_Summary": {"Coverage_Status"},
    "TC_to_Req_Map": {"Coverage_Status"},
    "Cross_Mismatch": {"Status"},
}

# ---- Classification keywords used during parsing/match
JUSTIFICATION_KEYWORDS = [
    "justification", "justified", "rationale", "reason"
]
OBJECT_TYPE_KEYS = ["object type", "type:"]
DERIVED_KEYWORDS = ["derived", "derivation"]

# ---------------------------------------------------------
# KEYWORDS
# ---------------------------------------------------------

# Used to detect "Object Type: Derived" requirements
OBJECT_TYPE_KEYS = [
    "object type"
]

# Used to detect derived requirements
DERIVED_KEYWORDS = [
    "derived"
]

# Used to detect justification text in traceability
JUSTIFICATION_KEYWORDS = [
    "justification"
]


# ---------------------------------------------------------
# STATUS LABELS
# ---------------------------------------------------------

STATUS_PASS = "PASS"
STATUS_FAIL = "FAIL"
STATUS_ERROR = "ERROR"
STATUS_NA = "N/A"


# ---------------------------------------------------------
# COMMON MESSAGES – COVERAGE & VALIDATION
# ---------------------------------------------------------

MSG_SRS_NOT_TRACEABLE = (
    "No mapping found in SRS→SYS traceability document"
)

MSG_SYS_NOT_TRACEABLE = (
    "No mapping found in SYS→SRS traceability document"
)

MSG_INVALID_SRS_ID = (
    "SRS ID used in traceability but not found in original SRS document"
)

MSG_INVALID_SYS_ID = (
    "SYS ID used in traceability but not found in original SYS document"
)


# ---------------------------------------------------------
# DERIVED REQUIREMENT MESSAGES (G6.3)
# ---------------------------------------------------------

MSG_DERIVED_OBJECT_TYPE = (
    "Derived Requirement (Object Type: Derived)"
)

MSG_DERIVED_JUSTIFIED = (
    "Derived Requirement (Justification Provided)"
)

MSG_POSSIBLE_DERIVED_OR_ERROR = (
    "ERROR – Possible Derived Requirement OR Missing SYS ID"
)


# ---------------------------------------------------------
# BIDIRECTIONAL TRACEABILITY (G6.4)
# ---------------------------------------------------------

MSG_BIDIRECTIONAL_MATCH = (
    "MATCH – Mapping exists in both SRS→SYS and SYS→SRS traceability documents."
)

MSG_BIDIRECTIONAL_MISMATCH = (
    "MISMATCH – Mapping exists in one direction but not the other."
)

MSG_DUPLICATE_MAPPING = (
    "DUPLICATE – Mapping already appeared earlier in the traceability document."
)

MSG_MISSING_SRS_ID = (
    "ERROR – SRS ID Missing in traceability document"
)

MSG_MISSING_SYS_ID = (
    "ERROR – SYS ID Missing in traceability document"
)


# ---------------------------------------------------------
# TEXT CONSISTENCY (G6.6)
# ---------------------------------------------------------

MSG_TEXT_EXACT_MATCH = "Exact Match"
MSG_TEXT_MISMATCH = "MisMatch"
MSG_TEXT_TABULAR = "N/A – Tabular requirement (ARINC table-based)"


# ---------------------------------------------------------
# MASTER SUMMARY DEFAULTS
# ---------------------------------------------------------

MASTER_PASS = "PASS"
MASTER_FAIL = "FAIL"
MASTER_NA = "N/A"

