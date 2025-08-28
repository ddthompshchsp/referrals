# services_referrals_app.py
import io
import re
from pathlib import Path
from datetime import datetime
from zoneinfo import ZoneInfo

import pandas as pd
import streamlit as st

st.set_page_config(page_title="HCHSP Services & Referrals", layout="wide")

# =========================
# Header
# =========================
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

# =========================
# Inputs
# =========================
inp_l, inp_c, inp_r = st.columns([1, 2, 1])
with inp_c:
    sref_file = st.file_uploader("Upload *10433.xlsx*", type=["xlsx"], key="sref")
    process = st.button("Process & Download")

# =========================
# Helpers
# =========================
def _norm(s: str) -> str:
    return re.sub(r"\s+", " ", str(s).strip().lower())

def detect_header_row(df_raw: pd.DataFrame) -> int:
    """Heuristic: row with many ST: cells and header-like keywords."""
    nrows = len(df_raw)
    best_idx, best_score = 0, -1
    keywords = ["date", "service", "result", "provided label", "detailed", "author", "staff", "user", "center", "name"]
    for i in range(min(nrows, 60)):
        vals = [str(v) for v in df_raw.iloc[i].tolist()]
        st_like = sum(1 for v in vals if v.strip().startswith("ST:"))
        kw_score = sum(1 for v in vals if any(k in v.lower() for k in keywords))
        score = st_like * 2 + kw_score
        if score > best_score:
            best_score, best_idx = score, i
    return best_idx

def clean_header(h: str) -> str:
    s = str(h).strip()
    s = re.sub(r"^(ST:|FD:)\s*", "", s, flags=re.I)
    s = re.sub(r"\s+", " ", s)
    return s

def uniqueize(names: list[str]) -> list[str]:
    seen, out = {}, []
    for n in names:
        if n not in seen:
            seen[n] = 1
            out.append(n)
        else:
            seen[n] += 1
            out.append(f"{n} ({seen[n]})")
    return out

def read_10433_keep_all(df_raw: pd.DataFrame) -> pd.DataFrame:
    """Read entire sheet, keep all columns, clean headers, drop fully-empty rows."""
    hdr_i = detect_header_row(df_raw)
    headers = df_raw.iloc[hdr_i].tolist()
    body = pd.DataFrame(df_raw.iloc[hdr_i + 1:].values, columns=headers)

    def row_empty(r: pd.Series) -> bool:
        s = r.astype(str).str.strip().replace({"nan": "", "NaT": ""})
        return r.isna().all() or s.eq("").all()

    body = body[~body.apply(row_empty, axis=1)].reset_index(drop=True)
    cleaned = [clean_header(c) for c in body.columns]
    body.columns = uniqueize(cleaned)
    return body

def parse_to_dt(series: pd.Series) -> pd.Series:
    """
    Parse strings + Excel serials into datetime (NaT on failure).
    - Strings: pd.to_datetime(..., errors='coerce')
    - Serials: plausible numbers as days since 1899-12-30
    """
    dt1 = pd.to_datetime(series, errors="coerce", infer_datetime_format=True)
    num = pd.to_numeric(series, errors="coerce")
    serial_mask = num.notna() & num.between(10000, 70000)  # broad window
    dt2 = pd.Series(pd.NaT, index=series.index, dtype="datetime64[ns]")
    if serial_mask.any():
        dt2.loc[serial_mask] = pd.to_datetime(num.loc[serial_mask], unit="D", origin="1899-12-30", errors="coerce")
    dt = dt1.copy()
    dt[dt.isna()] = dt2[dt.isna()]
    return dt

def find_service_date_col(df: pd.DataFrame) -> str | None:
    """Find 'Service Date' (post-clean)."""
    exact = {"service date", "date of service"}
    contains = ["service date", "date of service", "svc date", "service provided date", "date provided"]
    for c in df.columns:
        if _norm(c) in exact:
            return c
    for c in df.columns:
        cn = _norm(c)
        if any(k in cn for k in contains):
            return c
    for c in df.columns:
        cn = _norm(c)
        if "service" in cn and "date" in cn:
            return c
    return None

def filter_strict_by_service_date(df: pd.DataFrame, cutoff="2025-08-01") -> pd.DataFrame:
    """
    KEEP ONLY rows where Service Date is parseable AND >= cutoff.
    Drop earlier, missing, or unparseable Service Dates.
    Raise if the Service Date column can’t be found.
    """
    svc_col = find_service_date_col(df)
    if svc_col is None:
        raise ValueError("Could not find a 'Service Date' column (or close variant).")
    dt = parse_to_dt(df[svc_col])
    keep = dt.notna() & (dt >= pd.Timestamp(cutoff))
    return df.loc[keep].reset_index(drop=True)

def looks_like_date_col(series: pd.Series, header: str) -> bool:
    if "date" in _norm(header):
        return True
    dt = parse_to_dt(series)
    return dt.notna().mean() >= 0.5

def format_dates_for_preview(df: pd.DataFrame) -> pd.DataFrame:
    """Preview-only: show date-like columns as mm/dd/yy; leave non-parsable values unchanged."""
    out = df.copy()
    for col in out.columns:
        if looks_like_date_col(out[col], col):
            dt = parse_to_dt(out[col])
            mask = dt.notna()
            out.loc[mask, col] = dt[mask].dt.strftime("%m/%d/%y")
    return out

def col_letter(idx: int) -> str:
    s = ""
    n = idx
    while n >= 0:
        s = chr(n % 26 + 65) + s
        n = n // 26 - 1
    return s

def find_general_service_idx(df: pd.DataFrame) -> int:
    """Prefer a 'General Service' column; otherwise fallback to the densest text column."""
    for i, c in enumerate(df.columns):
        if _norm(c) == "general service":
            return i
    for i, c in enumerate(df.columns):
        if "general service" in _norm(c) or "provided label" in _norm(c):
            return i
    # fallback to densest non-empty text-like column
    best_i, best_score = 0, -1
    for i in range(df.shape[1]):
        s = df.iloc[:, i].astype(str).str.strip()
        score = int((s.ne("")) & (~s.isin(["nan", "NaT"]))).sum()
        if score > best_score:
            best_i, best_score = i, score
    return best_i

# =========================
# Excel writer (with filter-aware Total)
# =========================
def to_excel_styled(df: pd.DataFrame) -> bytes:
    """
    Styled Excel:
      - Logo at B1
      - Title + subtitle with CT date/time
      - Blue header; borders; thick outer frame
      - ONE dynamic Total at bottom: counts VISIBLE rows via hidden helper column + SUBTOTAL(109)
      - No Excel date formatting/coercion (write raw)
    """
    buf = io.BytesIO()
    with pd.ExcelWriter(buf, engine="xlsxwriter") as writer:
        sheet = "Head Start Services & Referrals"
        df.to_excel(writer, index=False, sheet_name=sheet, startrow=3)
        wb = writer.book
        ws = writer.sheets[sheet]

        # Layout basics
        ws.hide_gridlines(0)
        ws.set_row(0, 24); ws.set_row(1, 22); ws.set_row(2, 20)

        # Logo
        if logo_path.exists():
            ws.set_column(1, 1, 6)
            ws.insert_image(0, 1, str(logo_path), {
                "x_offset": 2, "y_offset": 2, "x_scale": 0.53, "y_scale": 0.53, "object_position": 1
            })

        # Titles with CT time
        now_ct = datetime.now(ZoneInfo("America/Chicago"))
        date_str = now_ct.strftime("%m/%d/%y %I:%M %p CT")
        title_fmt    = wb.add_format({"bold": True, "font_size": 14, "align": "center"})
        subtitle_fmt = wb.add_format({"bold": True, "font_size": 12, "align": "center"})
        red_fmt      = wb.add_format({"bold": True, "font_size": 12, "font_color": "#C00000"})

        last_col_0 = df.shape[1] - 1
        last_col_letter = col_letter(last_col_0)

        ws.merge_range(0, 2, 0, last_col_0, "Hidalgo County Head Start Program", title_fmt)
        ws.merge_range(1, 2, 1, last_col_0, "", subtitle_fmt)
        ws.write_rich_string(
            1, 2,
            subtitle_fmt, "Head Start - 2025-2026 Services and Referrals as of ",
            red_fmt, f"({date_str})",
            subtitle_fmt
        )

        # Header (blue)
        hdr_fmt = wb.add_format({
            "bold": True, "font_color": "white", "bg_color": "#305496",
            "align": "center", "valign": "vcenter", "text_wrap": True, "border": 1
        })
        ws.set_row(3, 26)
        for c, col in enumerate(df.columns):
            ws.write(3, c, col, hdr_fmt)

        # Data region
        last_row_0 = len(df) + 3           # 0-based last data row index
        last_excel_row = last_row_0 + 1    # Excel row number
        data_first_row0 = 4                # 0-based first data row (A5)
        data_first_excel_row = 5           # Excel row number of first data row

        # Autofilter over visible data (not including helper)
        ws.autofilter(3, 0, last_row_0, last_col_0)

        # Borders on table
        border_all = wb.add_format({"border": 1})
        ws.conditional_format(f"A4:{last_col_letter}{last_excel_row}",
                              {"type": "formula", "criteria": "TRUE", "format": border_all})

        # Column widths (simple defaults; no date formats)
        default_w = 20
        width_overrides = {
            "date": 14, "general service": 30, "detailed service": 36,
            "service author": 22, "result": 20, "center": 26,
            "participant": 26, "name": 26, "id": 16
        }
        for idx, name in enumerate(df.columns):
            key = _norm(name)
            width = next((w for k, w in width_overrides.items() if k in key), default_w)
            ws.set_column(idx, idx, width)

        # ===== Hidden helper column for filter-aware row count =====
        helper_idx = last_col_0 + 1
        helper_letter = col_letter(helper_idx)
        ws.write(3, helper_idx, "_helper_")  # header to avoid Excel nag
        for r0 in range(data_first_row0, last_row_0 + 1):
            ws.write_number(r0, helper_idx, 1)
        ws.set_column(helper_idx, helper_idx, None, None, {"hidden": 1})

        # ===== Total row (label + dynamic visible row count) =====
        totals_row_0 = last_row_0 + 1
        totals_excel_row = last_excel_row + 1
        label_fmt = wb.add_format({"bold": True, "align": "right"})
        value_fmt = wb.add_format({"bold": True, "align": "center"})

        ws.write(totals_row_0, 0, "Total", label_fmt)

        # Place the number under 'General Service' if present; otherwise under the first column
        gs_idx = find_general_service_idx(df)
        ws.write_formula(
            totals_row_0, gs_idx,
            f"=SUBTOTAL(109,{helper_letter}{data_first_excel_row}:{helper_letter}{last_excel_row})",
            value_fmt
        )

        # Border around totals row (visible area only)
        ws.conditional_format(f"A{totals_excel_row}:{last_col_letter}{totals_excel_row}",
                              {"type": "formula", "criteria": "TRUE", "format": border_all})

        # Thick outer box around title + visible table (not the hidden helper)
        top = wb.add_format({"top": 2}); bottom = wb.add_format({"bottom": 2})
        left = wb.add_format({"left": 2}); right = wb.add_format({"right": 2})
        ws.conditional_format(f"A1:{last_col_letter}1", {"type": "formula", "criteria": "TRUE", "format": top})
        ws.conditional_format(f"A1:A{totals_excel_row}", {"type": "formula", "criteria": "TRUE", "format": left})
        ws.conditional_format(f"{last_col_letter}1:{last_col_letter}{totals_excel_row}",
                              {"type": "formula", "criteria": "TRUE", "format": right})
        ws.conditional_format(f"A{totals_excel_row}:{last_col_letter}{totals_excel_row}",
                              {"type": "formula", "criteria": "TRUE", "format": bottom})

        # Ensure right edge on title rows
        ws.write(0, last_col_0, "", wb.add_format({"right": 2, "top": 2}))
        ws.write(1, last_col_0, "", wb.add_format({"right": 2}))

    return buf.getvalue()

# =========================
# Main
# =========================
if process and sref_file:
    try:
        raw = pd.read_excel(sref_file, sheet_name=0, header=None)
        df = read_10433_keep_all(raw)

        # STRICT service-date filter: ONLY keep rows with parseable Service Date >= 08/01/2025
        df = filter_strict_by_service_date(df, cutoff="2025-08-01")

        # Preview: date-like cols as mm/dd/yy (Excel gets raw values, no formatting)
        prev = format_dates_for_preview(df)

        st.success("Preview below. Use the download button to get the Excel file.")
        st.dataframe(prev, use_container_width=True)

        xlsx = to_excel_styled(df)
        st.download_button(
            "Download Services & Referrals (Excel)",
            data=xlsx,
            file_name="HCHSP_Services_Referrals_Formatted.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        )
    except Exception as e:
        st.error(f"Processing error: {e}")

