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
# Column detection helpers
# ----------------------------
def _normalize(s: str) -> str:
    return re.sub(r"\s+", " ", s.strip().lower())

def find_col(headers, candidates):
    """
    Find the first header whose normalized text contains any candidate
    (tries exact match first). Returns actual column name or None.
    """
    norm_map = {h: _normalize(str(h)) for h in headers}
    cand_norm = [_normalize(c) for c in candidates]
    # exact
    for c in cand_norm:
        for h, n in norm_map.items():
            if n == c:
                return h
    # contains
    for c in cand_norm:
        for h, n in norm_map.items():
            if c in n:
                return h
    return None

def detect_header_row(df_raw: pd.DataFrame) -> int:
    """
    Auto-detect header row by looking for ST:-style fields and common labels.
    """
    nrows = len(df_raw)
    best_idx, best_score = 0, -1
    keywords = ["date", "service", "result", "provided label", "detailed", "author", "staff", "user"]
    for i in range(min(nrows, 40)):
        row_vals = [str(v) for v in df_raw.iloc[i].tolist()]
        st_like = sum(1 for v in row_vals if isinstance(v, str) and v.strip().startswith("ST:"))
        kw_score = sum(1 for v in row_vals if any(k in str(v).lower() for k in keywords))
        score = st_like * 2 + kw_score
        if score > best_score:
            best_score, best_idx = score, i
    return best_idx

# ----------------------------
# Parser (10433)
# ----------------------------
def parse_services_referrals(df_raw: pd.DataFrame) -> pd.DataFrame:
    """
    Builds a tidy table with columns:
      Date | General Service | Detailed Service | Service Author | Result
    """
    header_row = detect_header_row(df_raw)
    headers = df_raw.iloc[header_row].tolist()
    body = pd.DataFrame(df_raw.iloc[header_row + 1:].values, columns=headers)

    # Candidate names for each output column
    date_cands = [
        "ST: Date", "ST: Service Date", "ST: Contact Date", "Date", "Service Date", "Provided Date"
    ]
    gen_service_cands = [
        "ST: General Service", "General Service", "Provided Label", "Service", "Service (General)"
    ]
    det_service_cands = [
        "ST: Detailed Service", "Detailed Service", "Detail Service", "Service (Detail)", "Service Detail"
    ]
    author_cands = [
        "ST: Service Author", "Service Author", "Provided By", "Staff", "Staff Name", "User", "ST: User", "ST: Staff Member"
    ]
    result_cands = [
        "ST: Result", "Result", "Service Result", "Outcome"
    ]

    date_col   = find_col(body.columns, date_cands)
    gen_col    = find_col(body.columns, gen_service_cands)
    det_col    = find_col(body.columns, det_service_cands)
    author_col = find_col(body.columns, author_cands)
    res_col    = find_col(body.columns, result_cands)

    missing = [("Date", date_col), ("General Service", gen_col),
               ("Detailed Service", det_col), ("Service Author", author_col), ("Result", res_col)]
    missing_names = [name for name, col in missing if col is None]
    if missing_names:
        raise ValueError(
            "Could not find required column(s): " + ", ".join(missing_names) +
            ". Please confirm 10433 export headers or share a sample row."
        )

    out = body[[date_col, gen_col, det_col, author_col, res_col]].copy()
    out.columns = ["Date", "General Service", "Detailed Service", "Service Author", "Result"]

    # Date coercion (safe)
    try:
        out["Date"] = pd.to_datetime(out["Date"], errors="coerce")
    except Exception:
        pass

    # Drop completely empty rows
    all_empty = out.apply(lambda r: r.isna().all() or (r.astype(str).str.strip() == "").all(), axis=1)
    out = out[~all_empty].reset_index(drop=True)

    # Display-friendly Date
    if np.issubdtype(out["Date"].dtype, np.datetime64):
        out["Date"] = out["Date"].dt.date

    return out

# ----------------------------
# Excel Writer (same styling + dynamic totals)
# ----------------------------
def to_styled_excel(df: pd.DataFrame) -> bytes:
    """
    Same aesthetics + dynamic totals row at the bottom (visible rows after filter).
    """
    def col_letter(n: int) -> str:
        s = ""
        while n >= 0:
            s = chr(n % 26 + 65) + s
            n = n // 26 - 1
        return s

    output = io.BytesIO()
    with pd.ExcelWriter(output, engine="xlsxwriter") as writer:
        sheet_name = "Head Start Services & Referrals"
        df.to_excel(writer, index=False, sheet_name=sheet_name, startrow=3)
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

        last_col_0 = len(df.columns) - 1
        last_col_letter = col_letter(last_col_0)

        # Main header line (brand)
        ws.merge_range(0, 2, 0, last_col_0, "Hidalgo County Head Start Program", title_fmt)
        # Subtitle per your request
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
        for c, col in enumerate(df.columns):
            ws.write(3, c, col, header_fmt)

        # Geometry
        last_row_0 = len(df) + 3               # 0-based index of last data row
        last_excel_row = last_row_0 + 1        # 1-based Excel row number
        data_first_excel_row = 5               # A5 is first data row (after header in A4)

        # Filters (no freeze panes)
        ws.autofilter(3, 0, last_row_0, last_col_0)

        # Column widths
        default_widths = {
            "Date": 14,
            "General Service": 30,
            "Detailed Service": 36,
            "Service Author": 22,
            "Result": 20,
        }
        for name, width in default_widths.items():
            if name in df.columns:
                idx = df.columns.get_loc(name)
                ws.set_column(idx, idx, width)

        # Borders on all header+data cells
        border_all = wb.add_format({"border": 1})
        ws.conditional_format(f"A4:{last_col_letter}{last_excel_row}", {
            "type": "formula", "criteria": "TRUE", "format": border_all
        })

        # ---- Dynamic totals row (counts visible rows after filter) ----
        totals_label_fmt = wb.add_format({"bold": True, "align": "right"})
        totals_val_fmt   = wb.add_format({"bold": True, "align": "center"})
        totals_row_0     = last_row_0 + 1          # 0-based row index for totals
        totals_excel_row = last_excel_row + 1      # 1-based Excel row number for totals

        # Label in first column
        ws.write(totals_row_0, 0, "Visible Rows (after filter):", totals_label_fmt)

        # Count visible using SUBTOTAL on a robust non-empty column (General Service)
        gs_idx = df.columns.get_loc("General Service")
        gs_letter = col_letter(gs_idx)
        data_range = f"{gs_letter}{data_first_excel_row}:{gs_letter}{last_excel_row}"
        ws.write_formula(totals_row_0, gs_idx, f"=SUBTOTAL(103,{data_range})", totals_val_fmt)

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
        tidy = parse_services_referrals(raw)

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


