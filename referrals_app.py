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
    """Auto-detect header row by looking for ST:-style fields and common labels."""
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
    """Strip 'ST:' or 'FD:' prefixes and collapse whitespace."""
    s = str(h).strip()
    s = re.sub(r"^(ST:|FD:)\s*", "", s, flags=re.I)
    s = re.sub(r"\s+", " ", s)
    return s

def make_unique(names: list[str]) -> list[str]:
    """Ensure column names are unique after cleaning."""
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
    """
    Preserve ALL columns from the 10433 export:
    - Detect header row
    - Clean headers (remove ST:/FD:) and ensure uniqueness
    - Drop completely empty rows
    """
    header_row = detect_header_row(df_raw)
    raw_headers = df_raw.iloc[header_row].tolist()
    body = pd.DataFrame(df_raw.iloc[header_row + 1:].values, columns=raw_headers)

    # Drop rows that are entirely empty/blank
    def _row_is_empty(r: pd.Series) -> bool:
        s = r.astype(str).str.strip().replace({"nan": "", "NaT": ""})
        return r.isna().all() or s.eq("").all()

    body = body[~body.apply(_row_is_empty, axis=1)].reset_index(drop=True)

    # Clean headers: strip ST:/FD: and ensure uniqueness
    cleaned = [clean_header_name(c) for c in body.columns]
    body.columns = make_unique(cleaned)
    return body

def _col_letter(n: int) -> str:
    """0-based index to Excel column letters."""
    s = ""
    while n >= 0:
        s = chr(n % 26 + 65) + s
        n = n // 26 - 1
    return s

def pick_count_column_index(df: pd.DataFrame) -> int:
    """Pick a robust column to use if we can't find General Service."""
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
    """
    Try to locate the 'General Service' column (post-cleaning).
    Checks common variants: 'General Service', 'Provided Label', 'Service' (but not 'Detailed'/'Type'/'Result').
    """
    candidates_exact = {"general service"}
    candidates_contains = ["general service", "provided label"]
    # exact
    for i, c in enumerate(df.columns):
        if _normalize(c) in candidates_exact:
            return i
    # contains
    for i, c in enumerate(df.columns):
        cn = _normalize(c)
        if any(k in cn for k in candidates_contains):
            return i
    # broad 'service' but not detailed/type/result
    for i, c in enumerate(df.columns):
        cn = _normalize(c)
        if "service" in cn and all(excl not in cn for excl in ["detailed", "type", "result"]):
            return i
    return None

def is_probably_date_series(s: pd.Series, header: str) -> bool:
    """Heuristic for date-like columns (by header or by parsability)."""
    name_has_date = "date" in _normalize(header)
    parsed = pd.to_datetime(s, errors="coerce", infer_datetime_format=True)
    has_any = parsed.notna().sum() > 0
    ratio = (parsed.notna().mean() if len(parsed) else 0.0)
    return name_has_date or (has_any and ratio >= 0.5)

def _as_datetime_for_preview(series: pd.Series, header: str) -> pd.Series:
    """
    Convert mixed date inputs into pandas datetime for preview only.
    Tries string dates first; then Excel serial numbers (origin 1899-12-30).
    """
    # Try normal parsing (strings like '2025-01-31', '01/31/25', etc.)
    dt1 = pd.to_datetime(series, errors="coerce", infer_datetime_format=True)

    # Try Excel serials (e.g., 45231)
    num = pd.to_numeric(series, errors="coerce")
    serial_mask = num.notna() & num.between(20000, 60000)  # rough modern serial range
    dt2 = pd.Series(pd.NaT, index=series.index, dtype="datetime64[ns]")
    if serial_mask.any():
        dt2.loc[serial_mask] = pd.to_datetime(
            num.loc[serial_mask], unit="D", origin="1899-12-30", errors="coerce"
        )

    # Prefer dt1; fill gaps with dt2
    dt = dt1.copy()
    dt[dt.isna()] = dt2[dt.isna()]
    return dt

def format_dates_for_preview(df: pd.DataFrame) -> pd.DataFrame:
    """
    Copy of df where date-like columns (including Excel serial dates) are shown as mm/dd/yy strings.
    Affects the Streamlit preview only; Excel export still writes raw values.
    """
    out = df.copy()
    for col in out.columns:
        if "date" in _normalize(col) or is_probably_date_series(out[col], header=col):
            dt = _as_datetime_for_preview(out[col], header=col)
            mask = dt.notna()
            out.loc[mask, col] = dt[mask].dt.strftime("%m/%d/%y")
    return out

def filter_out_before_cutoff_strict(df: pd.DataFrame, cutoff_str: str = "2025-08-01") -> pd.DataFrame:
    """
    STRICT: Drop any row where ANY cell in ANY column is a date < cutoff.
    - Parses strings like '8/1/25', '2025-08-01', etc. (coerce errors)
    - Interprets Excel serials (e.g., 45231) as days since 1899-12-30
    - Only dates within a plausible window [1900-01-01, 2100-12-31] are considered
    Rows with no parseable dates are kept.
    """
    if df.empty:
        return df

    cutoff = pd.Timestamp(cutoff_str)
    min_ok, max_ok = pd.Timestamp("1900-01-01"), pd.Timestamp("2100-12-31")

    drop_mask = pd.Series(False, index=df.index)

    for col in df.columns:
        s = df[col]

        # Try string-like dates
        dt_str = pd.to_datetime(s, errors="coerce", infer_datetime_format=True)
        valid_str = dt_str.notna() & (dt_str.between(min_ok, max_ok))
        drop_mask |= (valid_str & (dt_str < cutoff))

        # Try Excel serial numbers
        nums = pd.to_numeric(s, errors="coerce")
        serial_mask = nums.notna() & nums.between(10000, 70000)  # broaden range a bit
        if serial_mask.any():
            dt_serial = pd.to_datetime(nums, unit="D", origin="1899-12-30", errors="coerce")
            valid_serial = dt_serial.notna() & (dt_serial.between(min_ok, max_ok))
            drop_mask |= (valid_serial & (dt_serial < cutoff))

    return df.loc[~drop_mask].reset_index(drop=True)

# ----------------------------
# Excel Writer (no date auto-formatting; single dynamic total)
# ----------------------------
def to_styled_excel(df: pd.DataFrame) -> bytes:
    """
    - Same aesthetics: logo, title block, blue header, borders, thick outer box
    - NO Excel-side date coercion/formatting (writes values as-is)
    - ONE grand total only: count visible rows in the 'General Service' column (dynamic with filters)
      * Label cell reads 'Total'
      * Value cell uses SUBTOTAL(103, ...)
    """
    df_xls = df.copy()  # write exactly as provided

    output = io.BytesIO()
    with pd.ExcelWriter(output, engine="xlsxwriter") as writer:
        sheet_name = "Head Start Services & Referrals"
        df_xls.to_excel(writer, index=False, sheet_name=sheet_name, startrow=3)
        wb = writer.book
        ws = writer.sheets[sheet_name]

        # Keep Excel gridlines outside
        ws.hide_gridlines(0)

        # Title area heights
        ws.set_row(0, 24)
        ws.set_row(1, 22)
        ws.set_row(2, 20)

        # Logo at B1 (53% scale)
        logo = Path("header_logo.png")
        if logo.exists():
            ws.set_column(1, 1, 6)
            ws.insert_image(0, 1, str(logo), {
                "x_offset": 2, "y_offset": 2,
                "x_scale": 0.53, "y_scale": 0.53,
                "object_position": 1
            })

        # Titles across C..last (include Central time)
        now_ct = datetime.now(ZoneInfo("America/Chicago"))
        date_str = now_ct.strftime("%m/%d/%y %I:%M %p CT")

        title_fmt    = wb.add_format({"bold": True, "font_size": 14, "align": "center"})
        subtitle_fmt = wb.add_format({"bold": True, "font_size": 12, "align": "center"})
        red_fmt      = wb.add_format({"bold": True, "font_size": 12, "font_color": "#C00000"})

        last_col_0 = df_xls.shape[1] - 1
        last_col_letter = _col_letter(last_col_0)

        ws.merge_range(0, 2, 0, last_col_0, "Hidalgo County Head Start Program", title_fmt)
        ws.merge_range(1, 2, 1, last_col_0, "", subtitle_fmt)
        ws.write_rich_string(
            1, 2,
            subtitle_fmt, "Head Start - 2025-2026 Services and Referrals as of ",
            red_fmt, f"({date_str})",
            subtitle_fmt
        )

        # Header (blue)
        header_fmt = wb.add_format({
            "bold": True, "font_color": "white", "bg_color": "#305496",
            "align": "center", "valign": "vcenter", "text_wrap": True,
            "border": 1
        })
        ws.set_row(3, 26)
        for c, col in enumerate(df_xls.columns):
            ws.write(3, c, col, header_fmt)

        # Geometry
        last_row_0 = len(df_xls) + 3               # 0-based index of last data row
        last_excel_row = last_row_0 + 1            # 1-based Excel row number
        data_first_excel_row = 5                   # A5 is first data row (after header in A4)

        # Filters (no freeze panes)
        ws.autofilter(3, 0, last_row_0, last_col_0)

        # Column widths (no date formats applied)
        border_all = wb.add_format({"border": 1})
        default_width = 20
        width_overrides = {
            "date": 14,
            "general service": 30,
            "detailed service": 36,
            "service author": 22,
            "result": 20,
            "center": 26,
            "participant": 26,
            "name": 26,
            "id": 16,
        }
        for idx, name in enumerate(df_xls.columns):
            key = _normalize(name)
            width = default_width
            for k, w in width_overrides.items():
                if k in key:
                    width = w
                    break
            ws.set_column(idx, idx, width)

        # Borders on all header+data cells
        ws.conditional_format(f"A4:{last_col_letter}{last_excel_row}", {
            "type": "formula", "criteria": "TRUE", "format": border_all
        })

        # ---- Single dynamic TOTAL at the bottom (for General Service only) ----
        totals_label_fmt = wb.add_format({"bold": True, "align": "right"})
        totals_val_fmt   = wb.add_format({"bold": True, "align": "center"})
        totals_row_0     = last_row_0 + 1          # 0-based row index for totals
        totals_excel_row = last_excel_row + 1      # 1-based Excel row number for totals

        # Label in first column
        ws.write(totals_row_0, 0, "Total", totals_label_fmt)

        # Find the General Service column; fallback to robust count column if needed
        gs_idx = find_general_service_index(df_xls)
        if gs_idx is None:
            gs_idx = pick_count_column_index(df_xls)

        gs_letter = _col_letter(gs_idx)
        data_range = f"{gs_letter}{data_first_excel_row}:{gs_letter}{last_excel_row}"
        # Count visible, non-empty cells in General Service
        ws.write_formula(totals_row_0, gs_idx, f"=SUBTOTAL(103,{data_range})", totals_val_fmt)

        # Borders around totals row
        ws.conditional_format(f"A{totals_excel_row}:{last_col_letter}{totals_excel_row}", {
            "type": "formula", "criteria": "TRUE", "format": border_all
        })

        # ---- Thick outer box (row 1 → totals row). Fix right edge on title rows explicitly. ----
        top    = wb.add_format({"top": 2})
        bottom = wb.add_format({"bottom": 2})
        left   = wb.add_format({"left": 2})
        right  = wb.add_format({"right": 2})

        ws.conditional_format(f"A1:{last_col_letter}1", {"type": "formula", "criteria": "TRUE", "format": top})
        ws.conditional_format(f"A1:A{totals_excel_row}", {"type": "formula", "criteria": "TRUE", "format": left})
        ws.conditional_format(f"{last_col_letter}1:{last_col_letter}{totals_excel_row}", {"type": "formula", "criteria": "TRUE", "format": right})
        ws.conditional_format(f"A{totals_excel_row}:{last_col_letter}{totals_excel_row}", {"type": "formula", "criteria": "TRUE", "format": bottom})

        # Ensure right edge on title lines connects cleanly
        ws.write(0, last_col_0, "", wb.add_format({"right": 2, "top": 2}))
        ws.write(1, last_col_0, "", wb.add_format({"right": 2}))

    return output.getvalue()

# ----------------------------
# Main
# ----------------------------
if process and sref_file:
    try:
        raw = pd.read_excel(sref_file, sheet_name=0, header=None)
        tidy = parse_services_referrals_keep_all(raw)

        # STRICT FILTER: drop rows where any date in any column is < 08/01/2025
        tidy = filter_out_before_cutoff_strict(tidy, cutoff_str="2025-08-01")

        # Preview: display date-like columns as mm/dd/yy (handles Excel serials); Excel gets raw values
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


