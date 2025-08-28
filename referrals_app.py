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
# Header
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
# Helper Functions
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
# Strict filter by Service Date
# ----------------------------
def filter_by_service_date(df: pd.DataFrame, cutoff_str="2025-08-01") -> pd.DataFrame:
    df = df.copy()
    cutoff = pd.Timestamp(cutoff_str)

    # Find Service Date column
    service_date_col = None
    for c in df.columns:
        if "service date" in _normalize(c):
            service_date_col = c
            break
    if service_date_col is None:
        return pd.DataFrame(columns=df.columns)

    # Convert to datetime (string + Excel serial)
    dt_series = pd.to_datetime(df[service_date_col], errors="coerce", infer_datetime_format=True)
    nums = pd.to_numeric(df[service_date_col], errors="coerce")
    serial_mask = nums.notna() & nums.between(10000, 70000)
    dt_series[serial_mask] = pd.to_datetime(nums[serial_mask], unit="D", origin="1899-12-30", errors="coerce")

    # Keep rows where Service Date >= cutoff
    df = df.loc[dt_series >= cutoff].reset_index(drop=True)
    return df

# ----------------------------
# Excel export function (original full version)
# ----------------------------
def to_styled_excel(df: pd.DataFrame) -> bytes:
    df_xls = df.copy()
    output = io.BytesIO()
    with pd.ExcelWriter(output, engine="xlsxwriter") as writer:
        sheet_name = "Head Start Services & Referrals"
        df_xls.to_excel(writer, index=False, sheet_name=sheet_name, startrow=3)
        wb = writer.book
        ws = writer.sheets[sheet_name]

        ws.hide_gridlines(0)
        ws.set_row(0, 24)
        ws.set_row(1, 22)
        ws.set_row(2, 20)

        logo = Path("header_logo.png")
        if logo.exists():
            ws.set_column(1, 1, 6)
            ws.insert_image(0, 1, str(logo), {"x_offset":2, "y_offset":2, "x_scale":0.53, "y_scale":0.53, "object_position":1})

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

        header_fmt = wb.add_format({
            "bold": True, "font_color": "white", "bg_color": "#305496",
            "align": "center", "valign": "vcenter", "text_wrap": True,
            "border": 1
        })
        ws.set_row(3, 26)
        for c, col in enumerate(df_xls.columns):
            ws.write(3, c, col, header_fmt)

        last_row_0 = len(df_xls) + 3
        last_excel_row = last_row_0 + 1
        data_first_excel_row = 5

        ws.autofilter(3, 0, last_row_0, last_col_0)

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

        ws.conditional_format(f"A4:{last_col_letter}{last_excel_row}", {"type":"formula","criteria":"TRUE","format":border_all})

        totals_label_fmt = wb.add_format({"bold": True, "align": "right"})
        totals_val_fmt   = wb.add_format({"bold": True, "align": "center"})
        totals_row_0     = last_row_0 + 1
        totals_excel_row = last_excel_row + 1

        ws.write(totals_row_0, 0, "Total", totals_label_fmt)

        gs_idx = None
        for i, c in enumerate(df_xls.columns):
            if _normalize(c) in ["general service", "provided label"]:
                gs_idx = i
                break
        if gs_idx is None:
            gs_idx = 0

        gs_letter = _col_letter(gs_idx)
        data_range = f"{gs_letter}{data_first_excel_row}:{gs_letter}{last_excel_row}"
        ws.write_formula(totals_row_0, gs_idx, f"=SUBTOTAL(103,{data_range})", totals_val_fmt)

        ws.conditional_format(f"A{totals_excel_row}:{last_col_letter}{totals_excel_row}", {"type":"formula","criteria":"TRUE","format":border_all})

        # Thick outer box
        top    = wb.add_format({"top": 2})
        bottom = wb.add_format({"bottom": 2})
        left   = wb.add_format({"left": 2})
        right  = wb.add_format({"right": 2})

        ws.conditional_format(f"A1:{last_col_letter}1", {"type":"formula", "criteria":"TRUE", "format": top})
        ws.conditional_format(f"A1:A{totals_excel_row}", {"type":"formula", "criteria":"TRUE", "format": left})
        ws.conditional_format(f"{last_col_letter}1:{last_col_letter}{totals_excel_row}", {"type":"formula", "criteria":"TRUE", "format": right})
        ws.conditional_format(f"A{totals_excel_row}:{last_col_letter}{totals_excel_row}", {"type":"formula", "criteria":"TRUE", "format": bottom})

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

        # --- STRICT FILTER ---
        tidy = filter_by_service_date(tidy, cutoff_str="2025-08-01")

        # --- COLUMN ORDER ---
        column_order = [
            "Service Date", "PID", "First Name", "Last Name",
            "Center Name", "Class Name", "Author",
            "General Service", "Detailed Service", "Services Result", "Services - Result Date"
        ]
        for col in column_order:
            if col not in tidy.columns:
                tidy[col] = ""
        tidy = tidy[column_order]

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


