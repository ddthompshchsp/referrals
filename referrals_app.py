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
    seen = {}
    out = []
    for n in names:
        base = n
        if base not in seen:
            seen[base] = 1
            out.append(base)
        else:
            seen[base] += 1
            out.append(f"{base} ({seen[base]})")
    return out

def parse_services_referrals_keep_all(df_raw: pd.DataFrame) -> pd.DataFrame:
    """
    Build a tidy table that preserves ALL columns from the 10433 export.
    - Detects the most likely header row
    - Cleans headers (remove ST:/FD:) and keeps them unique
    - Drops completely empty rows
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
    cleaned = make_unique(cleaned)
    body.columns = cleaned

    return body

def _col_letter(n: int) -> str:
    """0-based index to Excel column letters."""
    s = ""
    while n >= 0:
        s = chr(n % 26 + 65) + s
        n = n // 26 - 1
    return s

def pick_count_column_index(df: pd.DataFrame) -> int:
    """Pick a robust column to use for visible-row counting with SUBTOTAL(103,...)."""
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

def is_probably_numeric_series(s: pd.Series) -> bool:
    """Heuristic for numeric columns."""
    coerced = pd.to_numeric(s, errors="coerce")
    if coerced.notna().sum() == 0:
        return False
    return coerced.notna().mean() >= 0.5

def is_probably_date_series(s: pd.Series, header: str) -> bool:
    """Heuristic for date-like columns (by header or by parsability)."""
    name_has_date = "date" in _normalize(header)
    parsed = pd.to_datetime(s, errors="coerce", infer_datetime_format=True)
    has_any = parsed.notna().sum() > 0
    ratio = (parsed.notna().mean() if len(parsed) else 0.0)
    return name_has_date or (has_any and ratio >= 0.5)

# ----------------------------
# Excel Writer (same styling + date formats + dynamic totals)
# ----------------------------
def to_styled_excel(df: pd.DataFrame) -> bytes:
    """
    - Same aesthetics: logo, title block, blue header, borders, thick outer box
    - Auto-format date-like columns to mm/dd/yy (display only)
    - Dynamic totals row:
        * SUBTOTAL(103, …) for visible row count in a robust text/date column
        * SUBTOTAL(109, …) to sum visible values for numeric columns
    """
    # Prepare a copy for export and convert date-like columns to datetime64
    df_xls = df.copy()
    date_like_idx = []
    for j, col in enumerate(df_xls.columns):
        if is_probably_date_series(df_xls[col], header=col):
            dt = pd.to_datetime(df_xls[col], errors="coerce", infer_datetime_format=True)
            if dt.notna().sum() > 0:
                df_xls[col] = dt
                date_like_idx.append(j)

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
            ws.set_column(1, 1, 6)  # column B width for logo
            ws.insert_image(0, 1, str(logo), {
                "x_offset": 2, "y_offset": 2,
                "x_scale": 0.53, "y_scale": 0.53,
                "object_position": 1
            })

        # Titles across C..last
        now_ct = datetime.now(ZoneInfo("America/Chicago"))
        date_str = now_ct.strftime("%m.%d.%y %I:%M %p CT")

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

        # Column widths + date formats
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
        date_fmt = wb.add_format({"num_format": "mm/dd/yy"})  # <- two-digit year
        for idx, name in enumerate(df_xls.columns):
            key = _normalize(name)
            width = default_width
            for k, w in width_overrides.items():
                if k in key:
                    width = w
                    break
            # Apply date format to detected date-like columns
            if idx in date_like_idx:
                ws.set_column(idx, idx, width, date_fmt)
            else:
                ws.set_column(idx, idx, width)

        # Borders on all header+data cells
        ws.conditional_format(f"A4:{last_col_letter}{last_excel_row}", {
            "type": "formula", "criteria": "TRUE", "format": border_all
        })

        # ---- Dynamic totals row (counts visible rows, sums numeric columns) ----
        totals_label_fmt = wb.add_format({"bold": True, "align": "right"})
        totals_val_fmt   = wb.add_format({"bold": True, "align": "center"})
        totals_row_0     = last_row_0 + 1          # 0-based row index for totals
        totals_excel_row = last_excel_row + 1      # 1-based Excel row number for totals

        # Label in first column
        ws.write(totals_row_0, 0, "Totals (visible):", totals_label_fmt)

        # Choose a robust column for counting visible rows
        count_idx = pick_count_column_index(df_xls)
        count_letter = _col_letter(count_idx)
        count_range = f"{count_letter}{data_first_excel_row}:{count_letter}{last_excel_row}"
        ws.write_formula(totals_row_0, count_idx, f"=SUBTOTAL(103,{count_range})", totals_val_fmt)

        # For every other column: if it's numeric-ish, write a SUBTOTAL(109, ...) to sum visible
        for j in range(df_xls.shape[1]):
            if j == count_idx:
                continue
            s = df_xls.iloc[:, j]
            if is_probably_numeric_series(s):
                col_letter = _col_letter(j)
                rng = f"{col_letter}{data_first_excel_row}:{col_letter}{last_excel_row}"
                ws.write_formula(totals_row_0, j, f"=SUBTOTAL(109,{rng})", totals_val_fmt)

        # Optional: draw borders around totals row too
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

        st.success("Preview below. Use the download button to get the Excel file.")
        st.dataframe(tidy, use_container_width=True)

        xlsx_bytes = to_styled_excel(tidy)
        st.download_button(
            "Download Services & Referrals (Excel)",
            data=xlsx_bytes,
            file_name="HCHSP_Services_Referrals_Formatted.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        )
    except Exception as e:
        st.error(f"Processing error: {e}")


