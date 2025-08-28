# services_referrals_app.py
import io
import re
from pathlib import Path
from datetime import datetime
from zoneinfo import ZoneInfo

import numpy as np
import pandas as pd
import streamlit as st

st.set_page_config(page_title="HCHSP Services & Referrals", layout="wide")

# ----------------------------
# Header (Streamlit UI only)
# ----------------------------
logo_path = Path("header_logo.png")
hdr_l, hdr_c, hdr_r = st.columns([1, 2, 1])
with hdr_c:
    if logo_path.exists():
        st.image(str(logo_path), width=320)
    st.markdown(
        "<h1 style='text-align:center; margin: 8px 0 4px;'>Hidalgo County Head Start — Services & Referrals</h1>",
        unsafe_allow_html=True,
    )
    st.markdown(
        """
        <p style='text-align:center; font-size:16px; margin-top:0;'>
        Upload the <strong>10433</strong> Services/Referrals report (Excel).
        </p>
        """,
        unsafe_allow_html=True,
    )

st.divider()

# ----------------------------
# Inputs
# ----------------------------
inp_l, inp_c, inp_r = st.columns([1, 2, 1])
with inp_c:
    sref_file = st.file_uploader("Upload *10433.xlsx*", type=["xlsx"], key="sref")
    process = st.button("Process & Download")

# ----------------------------
# Helpers
# ----------------------------
def _normalize(s: str) -> str:
    return re.sub(r"\s+", " ", str(s).strip().lower())

def detect_header_row(df_raw: pd.DataFrame) -> int:
    nrows = len(df_raw)
    best_idx, best_score = 0, -1
    keywords = ["date", "service", "result", "provided label", "detailed", "author", "staff", "user", "center", "name"]
    for i in range(min(nrows, 60)):
        row_vals = [str(v) for v in df_raw.iloc[i].tolist()]
        st_like = sum(1 for v in row_vals if isinstance(v, str) and v.strip().startswith("ST:"))
        kw_score = sum(1 for v in row_vals if any(k in str(v).lower() for k in keywords))
        score = st_like * 2 + kw_score
        if score > best_score:
            best_score, best_idx = score, i
    return best_idx

def clean_header_name(h: str) -> str:
    s = str(h).strip()
    s = re.sub(r"^(ST:|FD:)\s*", "", s, flags=re.I)
    s = re.sub(r"\s+", " ", s)
    return s

def make_unique(names: list[str]) -> list[str]:
    seen, out = {}, []
    for n in names:
        if n not in seen:
            seen[n] = 1
            out.append(n)
        else:
            seen[n] += 1
            out.append(f"{n} ({seen[n]})")
    return out

def parse_services_referrals_keep_all(df_raw: pd.DataFrame) -> pd.DataFrame:
    header_row = detect_header_row(df_raw)
    raw_headers = df_raw.iloc[header_row].tolist()
    body = pd.DataFrame(df_raw.iloc[header_row + 1:].values, columns=raw_headers)

    def _row_is_empty(r: pd.Series) -> bool:
        s = r.astype(str).str.strip().replace({"nan": "", "NaT": ""})
        return r.isna().all() or s.eq("").all()

    body = body[~body.apply(_row_is_empty, axis=1)].reset_index(drop=True)
    cleaned = [clean_header_name(c) for c in body.columns]
    body.columns = make_unique(cleaned)
    return body

def _col_letter(n: int) -> str:
    s = ""
    while n >= 0:
        s = chr(n % 26 + 65) + s
        n = n // 26 - 1
    return s

def pick_count_column_index(df: pd.DataFrame) -> int:
    best_idx, best_score, best_bonus = 0, -1, -1
    for i in range(df.shape[1]):
        s = df.iloc[:, i]
        s_str = s.astype(str).str.strip()
        nonempty = (s.notna()) & (s_str.ne("")) & (~s_str.isin(["nan", "NaT"]))
        score = int(nonempty.sum())
        bonus = 1 if s.dtype == object else 0
        if (score > best_score) or (score == best_score and bonus > best_bonus):
            best_idx, best_score, best_bonus = i, score, bonus
    return best_idx

def find_general_service_index(df: pd.DataFrame) -> int | None:
    candidates_exact = {"general service"}
    candidates_contains = ["general service", "provided label"]
    for i, c in enumerate(df.columns):
        if _normalize(c) in candidates_exact:
            return i
    for i, c in enumerate(df.columns):
        cn = _normalize(c)
        if any(k in cn for k in candidates_contains):
            return i
    for i, c in enumerate(df.columns):
        cn = _normalize(c)
        if "service" in cn and all(excl not in cn for excl in ["detailed", "type", "result"]):
            return i
    return None

def is_probably_date_series(s: pd.Series, header: str) -> bool:
    name_has_date = "date" in _normalize(header)
    parsed = pd.to_datetime(s, errors="coerce", infer_datetime_format=True)
    has_any = parsed.notna().sum() > 0
    ratio = (parsed.notna().mean() if len(parsed) else 0.0)
    return name_has_date or (has_any and ratio >= 0.5)

def _as_datetime_for_preview(series: pd.Series, header: str) -> pd.Series:
    dt1 = pd.to_datetime(series, errors="coerce", infer_datetime_format=True)
    num = pd.to_numeric(series, errors="coerce")
    serial_mask = num.notna() & num.between(20000, 60000)
    dt2 = pd.Series(pd.NaT, index=series.index, dtype="datetime64[ns]")
    if serial_mask.any():
        dt2.loc[serial_mask] = pd.to_datetime(
            num.loc[serial_mask], unit="D", origin="1899-12-30", errors="coerce"
        )
    dt = dt1.copy()
    dt[dt.isna()] = dt2[dt.isna()]
    return dt

def format_dates_for_preview(df: pd.DataFrame) -> pd.DataFrame:
    out = df.copy()
    for col in out.columns:
        if "date" in _normalize(col) or is_probably_date_series(out[col], header=col):
            dt = _as_datetime_for_preview(out[col], header=col)
            mask = dt.notna()
            out.loc[mask, col] = dt[mask].dt.strftime("%m/%d/%y")
    return out

# ----------------------------
# NEW: Strict cutoff filter
# ----------------------------
def filter_rows_after_cutoff(df: pd.DataFrame, cutoff_str="2025-08-01") -> pd.DataFrame:
    """
    Keep ONLY rows where all date-like columns are either >= cutoff or empty.
    Drop rows with any date < cutoff.
    """
    cutoff = pd.Timestamp(cutoff_str)
    df = df.copy()

    # Identify date-like columns heuristically
    date_cols = [c for c in df.columns if "date" in _normalize(c)]
    if not date_cols:
        date_cols = []
        for c in df.columns:
            dt = pd.to_datetime(df[c], errors="coerce", infer_datetime_format=True)
            if dt.notna().any():
                date_cols.append(c)

    if date_cols:
        for c in date_cols:
            dt_series = pd.to_datetime(df[c], errors="coerce", infer_datetime_format=True)
            nums = pd.to_numeric(df[c], errors="coerce")
            serial_mask = nums.notna() & nums.between(10000, 70000)
            dt_series[serial_mask] = pd.to_datetime(nums[serial_mask], unit="D", origin="1899-12-30", errors="coerce")
            df[c + "_dt"] = dt_series

        dt_cols = [c for c in df.columns if c.endswith("_dt")]
        mask_keep = ~(df[dt_cols] < cutoff).any(axis=1)
        df = df.loc[mask_keep].reset_index(drop=True)
        df = df[df.columns.difference(dt_cols)]

    return df

# ----------------------------
# Excel writer function (unchanged)
# ----------------------------
def to_styled_excel(df: pd.DataFrame) -> bytes:
    # ... (same code as your original to_styled_excel function)
    # For brevity, reuse the existing function from your code
    pass

# ----------------------------
# Main
# ----------------------------
if process and sref_file:
    try:
        raw = pd.read_excel(sref_file, sheet_name=0, header=None)
        tidy = parse_services_referrals_keep_all(raw)

        # --- STRICT FILTER: remove all rows with any date before 08/01/2025 ---
        tidy = filter_rows_after_cutoff(tidy, cutoff_str="2025-08-01")

        # Preview
        preview_df = format_dates_for_preview(tidy)
        st.success("Preview below. Use the download button to get the Excel file.")
        st.dataframe(preview_df, use_container_width=True)

        xlsx_bytes = to_styled_excel(tidy)
        st.download_button(
            "Download Services & Referrals (Excel)",
            data=xlsx_bytes,
            file_name="HCHSP_Services_Referrals_Formatted.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        )
    except Exception as e:
        st.error(f"Processing error: {e}")


