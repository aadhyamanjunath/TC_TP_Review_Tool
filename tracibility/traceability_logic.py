# traceability_logic.py
from __future__ import annotations
import re
from typing import List, Dict, Optional, Tuple
import pandas as pd
import config

# =========================
# ========= REGEX =========
# =========================
# Patterns used to find normalized SRS and TC IDs in free text
SRS_PAT = re.compile(
    r'(?:MRJ[\s_\-]*SCU[\s_\-]*STC[\s_\-]*)?SRS[\s_\-]*([0-9]+(?:[\s_\-]*_[\s_\-]*[0-9]+)*)',
    flags=re.IGNORECASE
)
ALT_SRS_PAT = re.compile(
    r'(?:MRJ[\s_\-]*SCU[\s_\-]*STC[\s_\-]*)?SRS[\s_\-]*([0-9]+(?:[\s_\-]*[ _\-][\s_\-]*[0-9]+)*)',
    flags=re.IGNORECASE
)
TC_PAT = re.compile(r'\bTC[-_A-Z0-9]*\d+\b', re.IGNORECASE)


# =========================
# ======== HELPERS ========
# =========================
def _collapse_delims(s: str) -> str:
    return re.sub(r'[_\-\s]+', '_', s.strip())


def normalize_req_ids(text: str) -> List[str]:
    if not isinstance(text, str):
        return []
    found = set()
    for m in SRS_PAT.finditer(text):
        body = _collapse_delims(m.group(1))
        found.add(f"SRS_{body.upper()}")
    if not found:
        for m in ALT_SRS_PAT.finditer(text):
            body = _collapse_delims(m.group(1))
            found.add(f"SRS_{body.upper()}")
    return sorted(found)


def find_tc_ids(text: str) -> List[str]:
    if not isinstance(text, str):
        return []
    return sorted({m.group(0).upper() for m in TC_PAT.finditer(text)})


def _format_cell(val) -> str:
    if pd.isna(val):
        return ""
    if isinstance(val, float):
        return str(int(val)) if val.is_integer() else f"{val:.6g}"
    return str(val)


def _normalize_dataframe(df: pd.DataFrame) -> pd.DataFrame:
    df = df.copy()
    df.columns = [str(c).strip() for c in df.columns]
    return df


def normalize_value(x, placeholders_lower: set):
    if pd.isna(x):
        return pd.NA
    s = str(x).strip()
    if s.lower() in placeholders_lower:
        return pd.NA
    return s


def resolve_column(df: pd.DataFrame, wanted: str) -> str:
    if wanted in df.columns:
        return wanted
    lower_map = {c.lower(): c for c in df.columns}
    key = wanted.lower()
    if key in lower_map:
        return lower_map[key]
    raise KeyError(f"Column '{wanted}' not found (available: {list(df.columns)})")


# =========================
# AUTO-DETECTION HELPERS
# =========================
def normalize_header(s: str) -> str:
    s = re.sub(r"[_\-:]+", " ", str(s).strip().lower())
    s = re.sub(r"[^a-z0-9]+", " ", s)
    s = re.sub(r"\s+", " ", s).strip()
    return s


def header_has_tokens(header_norm: str, tokens: List[str]) -> bool:
    words = set(header_norm.split())
    for t in tokens:
        t_norm = normalize_header(t)
        t_words = t_norm.split()
        if all(w in words for w in t_words):
            return True
    return False


def score_header_for_role(header_norm: str, role: str) -> int:
    score = 0
    if role == "req":
        if header_has_tokens(header_norm, config.REQ_HEADER_HINTS): score += 3
        if header_has_tokens(header_norm, config.ID_SYNONYMS):      score += 2
        if header_has_tokens(header_norm, config.REQ_NEGATIVE_HINTS): score -= 5
    elif role == "tc":
        if header_has_tokens(header_norm, config.TC_HEADER_HINTS):  score += 3
        if header_has_tokens(header_norm, config.ID_SYNONYMS):      score += 2
        if header_has_tokens(header_norm, config.TC_NEGATIVE_HINTS): score -= 5
    return score


def value_match_ratio(series: pd.Series, patterns: List[str], sample_size: int = 300) -> float:
    s = series.dropna()
    if s.empty:
        return 0.0
    s = s.astype(str).head(sample_size)
    regexes = [re.compile(p) for p in patterns]

    def matches_any(x: str) -> bool:
        xx = x.strip()
        for rgx in regexes:
            if rgx.fullmatch(xx):
                return True
        return False

    matched = s.map(matches_any).sum()
    total = len(s)
    return matched / total if total else 0.0


def detect_id_column(df: pd.DataFrame, role: str) -> str:
    role = role.lower().strip()
    if role not in {"req", "tc"}:
        raise ValueError("role must be 'req' or 'tc'")
    cols = list(df.columns)
    if not cols:
        raise KeyError("No columns in DataFrame")
    header_norm_map = {c: normalize_header(c) for c in cols}
    header_scores = {c: score_header_for_role(header_norm_map[c], role) for c in cols}
    candidates = [c for c in cols if header_scores[c] >= 0] or cols[:]
    patterns = config.REQ_ID_CANDIDATE_REGEXES if role == "req" else config.TC_ID_CANDIDATE_REGEXES
    value_scores = {}
    for c in candidates:
        try:
            ratio = value_match_ratio(df[c], patterns)
        except Exception:
            ratio = 0.0
        value_scores[c] = ratio
    combined = []
    for c in candidates:
        combined_score = header_scores.get(c, 0) * 1.0 + value_scores.get(c, 0.0) * 5.0
        combined.append((combined_score, value_scores.get(c, 0.0), header_scores.get(c, 0), c))
    combined.sort(key=lambda t: (t[0], t[1], t[2], -len(t[3])), reverse=True)
    best = combined[0] if combined else None
    if not best:
        raise KeyError(f"Could not auto-detect a '{role}' column")
    _, val_ratio, hdr_score, best_col = best
    if hdr_score >= 2 or val_ratio >= 0.20:
        return best_col
    canon_try = config.DEFAULT_REQ_COL if role == "req" else config.DEFAULT_TC_COL
    lower_map = {c.lower(): c for c in cols}
    if canon_try.lower() in lower_map:
        return lower_map[canon_try.lower()]
    raise KeyError(f"Could not confidently detect a '{role}' column")


def resolve_column_or_detect(df: pd.DataFrame, wanted: str, role: str) -> str:
    try:
        return resolve_column(df, wanted)
    except KeyError:
        return detect_id_column(df, role)


def find_header_row_and_cols_anywhere(df_matrix: pd.DataFrame,
                                      req_hdr_threshold: int = 2,
                                      tc_hdr_threshold: int = 2) -> Optional[Tuple[int, int, int]]:
    n_rows, n_cols = df_matrix.shape
    if n_rows == 0 or n_cols == 0:
        return None
    for r in range(n_rows):
        best_req = (-10, -1)
        best_tc = (-10, -1)
        for c in range(n_cols):
            val = df_matrix.iat[r, c]
            if pd.isna(val):
                continue
            h = normalize_header(str(val))
            if not h:
                continue
            req_score = score_header_for_role(h, "req")
            tc_score = score_header_for_role(h, "tc")
            if req_score > best_req[0]:
                best_req = (req_score, c)
            if tc_score > best_tc[0]:
                best_tc = (tc_score, c)
        if best_req[0] >= req_hdr_threshold and best_tc[0] >= tc_hdr_threshold and best_req[1] != best_tc[1]:
            return (r, best_req[1], best_tc[1])
    return None


def detect_columns_by_values_anywhere(df_matrix: pd.DataFrame) -> Optional[Tuple[int, int]]:
    n_rows, n_cols = df_matrix.shape
    if n_rows == 0 or n_cols == 0:
        return None
    req_scores = []
    tc_scores = []
    for c in range(n_cols):
        s = df_matrix.iloc[:, c]
        try:
            req_ratio = value_match_ratio(s.astype(str), config.REQ_ID_CANDIDATE_REGEXES)
        except Exception:
            req_ratio = 0.0
        try:
            tc_ratio = value_match_ratio(s.astype(str), config.TC_ID_CANDIDATE_REGEXES)
        except Exception:
            tc_ratio = 0.0
        req_scores.append((req_ratio, c))
        tc_scores.append((tc_ratio, c))
    req_scores.sort(reverse=True)
    tc_scores.sort(reverse=True)
    MIN_RATIO = 0.15
    best_req = req_scores[0] if req_scores else (0.0, -1)
    best_tc = tc_scores[0] if tc_scores else (0.0, -1)
    if best_req[1] == best_tc[1]:
        next_tc = tc_scores[1] if len(tc_scores) > 1 else (0.0, -1)
        if next_tc[0] >= MIN_RATIO:
            best_tc = next_tc
        else:
            next_req = req_scores[1] if len(req_scores) > 1 else (0.0, -1)
            if next_req[0] >= MIN_RATIO:
                best_req = next_req
    if best_req[0] >= MIN_RATIO and best_tc[0] >= MIN_RATIO and best_req[1] != best_tc[1]:
        return (best_req[1], best_tc[1])
    return None


def build_df_from_anywhere_scan(df_matrix: pd.DataFrame,
                                placeholders_lower: set,
                                canonical_req: str,
                                canonical_tc: str) -> Optional[pd.DataFrame]:
    header_found = find_header_row_and_cols_anywhere(df_matrix)
    if header_found:
        hdr_row, req_c, tc_c = header_found
        data = df_matrix.iloc[hdr_row+1:, [req_c, tc_c]].copy()
        header_req_text = str(df_matrix.iat[hdr_row, req_c]) if not pd.isna(df_matrix.iat[hdr_row, req_c]) else canonical_req
        header_tc_text = str(df_matrix.iat[hdr_row, tc_c]) if not pd.isna(df_matrix.iat[hdr_row, tc_c]) else canonical_tc
        data.columns = [header_req_text, header_tc_text]
        data = data.rename(columns={data.columns[0]: canonical_req, data.columns[1]: canonical_tc})
        data[canonical_req] = data[canonical_req].apply(lambda x: normalize_value(x, placeholders_lower))
        data[canonical_tc] = data[canonical_tc].apply(lambda x: normalize_value(x, placeholders_lower))
        data = data[~(data[canonical_req].isna() & data[canonical_tc].isna())].copy()
        return data
    detected = detect_columns_by_values_anywhere(df_matrix)
    if detected:
        req_c, tc_c = detected
        data = df_matrix.iloc[:, [req_c, tc_c]].copy()
        data.columns = [canonical_req, canonical_tc]
        data[canonical_req] = data[canonical_req].apply(lambda x: normalize_value(x, placeholders_lower))
        data[canonical_tc] = data[canonical_tc].apply(lambda x: normalize_value(x, placeholders_lower))
        data = data[~(data[canonical_req].isna() & data[canonical_tc].isna())].copy()
        return data
    return None


def split_cell_to_list(value: object, placeholders_lower: set,
                       strip_wrapping_brackets: bool, multi_value_separators: str) -> List[object]:
    if pd.isna(value):
        return [pd.NA]
    s = str(value).strip()
    if strip_wrapping_brackets:
        if (s.startswith("[") and s.endswith("]")) or (s.startswith("(") and s.endswith(")")):
            s = s[1:-1].strip()
    if s.lower() in placeholders_lower:
        return [pd.NA]
    parts = re.split(multi_value_separators, s)
    out = []
    for p in parts:
        p = p.strip()
        if not p:
            continue
        if p.lower() in placeholders_lower:
            continue
        out.append(p)
    return out if out else [pd.NA]


def explode_multivalue_rows(df: pd.DataFrame, req_col: str, tc_col: str,
                            placeholders_lower: set,
                            enable_split: bool, strip_wrapping_brackets: bool,
                            multi_value_separators: str,
                            passthrough_cols: List[str] = None) -> pd.DataFrame:
    if not enable_split:
        out = df.copy()
        if "_SourceRow" not in out.columns:
            out["_SourceRow"] = out.index + 2
        return out
    passthrough_cols = passthrough_cols or []
    records = []
    for idx, row in df.iterrows():
        req_list = split_cell_to_list(
            row[req_col], placeholders_lower, strip_wrapping_brackets, multi_value_separators
        )
        tc_list = split_cell_to_list(
            row[tc_col], placeholders_lower, strip_wrapping_brackets, multi_value_separators
        )
        if all(pd.isna(r) for r in req_list) and all(pd.isna(t) for t in tc_list):
            continue
        for r in req_list:
            for t in tc_list:
                rec = {
                    req_col: (pd.NA if pd.isna(r) else r),
                    tc_col: (pd.NA if pd.isna(t) else t),
                    "_SourceRow": idx + 2
                }
                for c in passthrough_cols:
                    rec[c] = row.get(c, pd.NA)
                records.append(rec)
    cols = [req_col, tc_col, "_SourceRow"] + passthrough_cols
    return pd.DataFrame.from_records(records, columns=cols) if records else pd.DataFrame(columns=cols)


def validate_id_patterns(series: pd.Series, pattern: str) -> pd.DataFrame:
    regex = re.compile(pattern)
    s = series.dropna().astype(str)
    ok = s.map(lambda x: bool(regex.fullmatch(x)))
    bad_series = s[~ok]
    return pd.DataFrame({"ID": bad_series.values})


def suggest_requirements(tc_id: str, all_requirements: list) -> str:
    if not isinstance(tc_id, str):
        tc_id = str(tc_id)
    tokens = re.findall(r"\d+", tc_id)
    if not tokens:
        return ""
    suggestions, seen = [], set()
    for req in all_requirements:
        sreq = str(req)
        for t in tokens:
            if t in sreq and sreq not in seen:
                suggestions.append(sreq); seen.add(sreq)
                break
        if len(suggestions) >= 10:
            break
    return ", ".join(suggestions)


def suggest_tcs(requirement_id: str, all_tcs: list) -> str:
    if not isinstance(requirement_id, str):
        requirement_id = str(requirement_id)
    tokens = re.findall(r"\d+", requirement_id)
    if not tokens:
        return ""
    suggestions, seen = [], set()
    for tc in all_tcs:
        stc = str(tc)
        for t in tokens:
            if t in stc and stc not in seen:
                suggestions.append(stc); seen.add(stc)
                break
        if len(suggestions) >= 10:
            break
    return ", ".join(suggestions)


def looks_like_id(token: str) -> bool:
    token = token.strip().strip(",.;:()[]")
    if " " in token or "\t" in token:
        return False
    return "_" in token and any(ch.isdigit() for ch in token) and any(ch.isalpha() for ch in token)


def is_object_type_derived(text: Optional[str]) -> bool:
    if not text:
        return False
    text_lower = text.lower()
    has_object_type = any(key in text_lower for key in config.OBJECT_TYPE_KEYS)
    has_derived = any(word in text_lower for word in config.DERIVED_KEYWORDS)
    return has_object_type and has_derived


def cross_match(s2tc_df: pd.DataFrame, tc2s_df: pd.DataFrame, req_df: pd.DataFrame) -> pd.DataFrame:
    valid_req_ids = set(req_df["ID"].astype(str))

    def normalize(a, b):
        if a is None or b is None:
            return None
        if "_req_" in a and "_tc_" in b: return (a, b)
        if "_tc_" in a and "_req_" in b: return (b, a)
        return None

    req2tc_pairs = set()
    for a, b in zip(s2tc_df["FROM_ID"], s2tc_df["TO_ID"]):
        pair = normalize(a, b)
        if pair:
            req2tc_pairs.add(pair)

    tc2s_pairs = set()
    for a, b in zip(tc2s_df["FROM_ID"], tc2s_df["TO_ID"]):
        pair = normalize(a, b)
        if pair:
            tc2s_pairs.add(pair)

    rows = []
    seen = set()
    seen_rev = set()

    for a, b, justify_flag, trace_text in zip(
        s2tc_df["FROM_ID"], s2tc_df["TO_ID"], s2tc_df["HAS_JUSTIFICATION"], s2tc_df["FROM_TEXT"]
    ):
        if a is None or str(a).strip() == "":
            rows.append([None, b if b else None, "ERROR – req ID Missing in req→tc Traceability Document"])
            continue
        if b is None or str(b).strip() == "":
            if is_object_type_derived(trace_text):
                rows.append([a, None, "Derived Requirement (Object Type: Derived)"])
                continue
            if justify_flag:
                rows.append([a, None, "Derived Requirement (Justification Provided)"])
                continue
            rows.append([a, None, "ERROR – Possible Derived Req OR tc ID Missing in req→tc Traceability Document"])
            continue
        if a not in valid_req_ids:
            rows.append([a, b if b else None, "ERROR – req ID not found in original req document"])
            continue
        norm = normalize(a, b)
        if not norm:
            continue
        req, tc = norm
        if norm in seen:
            status = "DUPLICATE – This req→tc mapping already appeared earlier in the traceability document."
        else:
            status = "MATCH – Mapping exists in both req→tc and tc→req traceability documents." if norm in tc2s_pairs else \
                     "MISMATCH – Mapping exists in req→tc but NOT found in tc→req traceability document."
            seen.add(norm)
        rows.append([req, tc, status])

    for a, b in zip(tc2s_df["FROM_ID"], tc2s_df["TO_ID"]):
        if a is None or str(a).strip() == "":
            rows.append([b if b else "None", None, "ERROR – tc ID Missing in tc→req Traceability Document"])
            continue
        if b is None or str(b).strip() == "":
            rows.append([None, a, "ERROR – req ID Missing in tc→req Traceability Document"])
            continue
        if b not in valid_req_ids:
            rows.append([b, a, "ERROR – req ID not found in original req document"])
            continue
        norm = normalize(a, b)
        if norm:
            if norm in seen_rev:
                rows.append([norm[1], norm[0], "DUPLICATE – This tc→req mapping already appeared earlier in the traceability document."])
                continue
            else:
                seen_rev.add(norm)
        if norm and norm not in req2tc_pairs:
            rows.append([norm[0], norm[1], "MISMATCH – Mapping exists in tc→req but NOT found in req→tc traceability document."])
            continue

    return pd.DataFrame(rows, columns=["req_ID", "tc_ID", "Status"]).drop_duplicates()


def check_req_trace_coverage(req_df: pd.DataFrame, trace_s2tc_df: pd.DataFrame) -> pd.DataFrame:
    original_req_ids = set(req_df["ID"].astype(str))
    traced_req_ids = set(trace_s2tc_df["FROM_ID"].dropna().astype(str))
    missing_in_trace = original_req_ids - traced_req_ids
    coverage_issues = pd.DataFrame(
        [
            [req_id, "req", "NOT TRACEABLE", "No mapping found in req→tc traceability document"]
            for req_id in missing_in_trace
        ],
        columns=["ID", "Document", "Issue_Type", "Details"]
    )
    return coverage_issues


def check_invalid_req_ids(trace_s2tc_df: pd.DataFrame, req_df: pd.DataFrame) -> pd.DataFrame:
    valid_req_ids = set(req_df["ID"].astype(str))
    traced_req_ids = set(trace_s2tc_df["FROM_ID"].dropna().astype(str))
    invalid_ids = traced_req_ids - valid_req_ids
    return pd.DataFrame(
        [
            [req_id, "req", "INVALID_ID", "req ID used in traceability but not found in original req document"]
            for req_id in invalid_ids
        ],
        columns=["ID", "Document", "Issue_Type", "Details"]
    )


def _join_unique(sorted_list: List[str]) -> str:
    return ", ".join(sorted_list)


def build_bidirectional_maps(matrix_df: pd.DataFrame) -> Tuple[pd.DataFrame, pd.DataFrame]:
    req_groups = (matrix_df
                  .drop_duplicates(subset=["requirement_id", "tc_id"])
                  .groupby("requirement_id")["tc_id"]
                  .apply(lambda s: sorted(s.unique()))
                  .reset_index(name="tc_ids"))
    req_groups["tc_count"] = req_groups["tc_ids"].apply(len)
    req_groups["tc_ids"] = req_groups["tc_ids"].apply(_join_unique)
    req_to_tc = req_groups.sort_values(["tc_count", "requirement_id"], ascending=[False, True])

    tc_groups = (matrix_df
                 .drop_duplicates(subset=["requirement_id", "tc_id"])
                 .groupby("tc_id")["requirement_id"]
                 .apply(lambda s: sorted(s.unique()))
                 .reset_index(name="requirement_ids"))
    tc_groups["requirement_count"] = tc_groups["requirement_ids"].apply(len)
    tc_groups["requirement_ids"] = tc_groups["requirement_ids"].apply(_join_unique)
    tc_to_req = tc_groups.sort_values(["requirement_count", "tc_id"], ascending=[False, True])
    return req_to_tc, tc_to_req


def build_master_summary(req_df: pd.DataFrame, cross_df: pd.DataFrame, coverage_df: pd.DataFrame) -> pd.DataFrame:
    rows = []
    req_ids = sorted(set(req_df["ID"].astype(str)) | set(cross_df["req_ID"].dropna().astype(str)))
    for req_id in req_ids:
        invalid_req_ids = set(
            cross_df[
                cross_df["Status"].str.contains(
                    "req ID not found in original req document",
                    na=False
                )
            ]["req_ID"]
        )
        if req_id in invalid_req_ids:
            g61 = "FAIL\nreq ID not found in original req document"
            has_trace = False
        elif req_id in coverage_df["ID"].values:
            g61 = "FAIL\nNo mapping found in req→tc traceability document"
            has_trace = False
        else:
            g61 = "PASS"
            has_trace = True
        if not has_trace:
            rows.append([req_id, g61, "N/A", "N/A", "FAIL", "FAIL"])
            continue
        g62 = "PASS"
        derived_rows = cross_df[(cross_df["req_ID"] == req_id) & (cross_df["Status"].str.contains("Derived Requirement \\(", na=False))]
        g63 = f"PASS\n{derived_rows.iloc[0]['Status']}" if not derived_rows.empty else "N/A\nNot a derived requirement"
        cross_rows = cross_df[cross_df["req_ID"] == req_id]
        g64 = "FAIL\nBidirectional mapping missing or mismatched" if any(cross_rows["Status"].str.contains("MISMATCH|ERROR", na=False)) else "PASS\nBidirectional traceability maintained"
        g65 = "FAIL\nInvalid or incomplete traceability entries detected" if any(cross_rows["Status"].str.contains("ERROR", na=False)) else "PASS\nTraceability matrix complete and accurate"
        rows.append([req_id, g61, g62, g63, g64, g65])
    return pd.DataFrame(rows, columns=["Requirement_ID", "6.1", "6.2", "6.3", "6.4", "6.5"])

def generate_excel_report(req_path, trace_s2tc_path, trace_tc2s_path, outfile):
    trace_s2tc_df = parse_traceability(trace_s2tc_path)
