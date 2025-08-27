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
    norm_map = {h: _normalize(str(h)) for h in headers}
    cand_norm = [_normalize(c) for c in candidates]
    for c in cand_norm:
        for h, n in norm_map.items():
            if n == c:
                return h
    for c in cand_norm:
        for h, n in norm_map.items():
            if c in n:
                return h
    return None

def detect_header_row(df_raw: pd.DataFrame) -> int:
    nrows = len(df_raw)
    best_idx, best_score = None, -1
    keywords = ["date", "pid", "first", "last", "center", "service", "result"]
    for i in range(min(nrows, 40)):
        row_vals = [str(v) for v in df_raw.iloc[i].tolist()]
        kw_score = sum(1 for v in row_vals if any(k in str(v).lower() for k in keywords))
        if kw_score > best_score:
            best_score, best_idx = kw_score, i
    return best_idx if best_idx is not None else 0

# ----------------------------
# Parser (10433)
# ----------------------------
def parse_services_referrals(df_raw: pd.DataFrame) -> pd.DataFrame:
    header_row = detect_header_row(df_raw)
    headers = df_raw.iloc[header_row].tolist()
    body = pd.DataFrame(df_raw.iloc[header_row + 1:].values, columns=headers)

    # Candidate names
    date_cands  = ["ST: Date", "Service Date", "Date"]
    pid_cands   = ["ST: Participant PID", "PID", "Participant PID"]
    fname_cands = ["ST: First Name", "First Name"]
    lname_cands = ["ST: Last Name", "Last Name"]
    center_cands= ["ST: Center Name", "Center", "Location"]

    gen_service_cands = ["General Service", "ST: General Service", "Provided Label"]
    det_service_cands = ["Detailed Service", "ST: Detailed Service", "Detail Service"]
    type_cands        = ["Service Type", "ST: Service Type", "Type"]
    result_cands      = ["Result", "ST: Result", "Outcome"]

    # Resolve
    date_col  = find_col(body.columns, date_cands)
    pid_col   = find_col(body.columns, pid_cands)
    fname_col = find_col(body.columns, fname_cands)
    lname_col = find_col(body.columns, lname_cands)
    center_col= find_col(body.columns, center_cands)
    gen_col   = find_col(body.columns, gen_service_cands)
    det_col   = find_col(body.columns, det_service_cands)
    type_col  = find_col(body.columns, type_cands)
    res_col   = find_col(body.columns, result_cands)

    cols = {
        "Date": date_col, "PID": pid_col, "First Name": fname_col, "Last Name": lname_col,
        "Center": center_col, "General Service": gen_col, "Detailed Service": det_col,
        "Service Type": type_col, "Result": res_col
    }
    missing = [k for k,v in cols.items() if v is None]
    if missing:
        raise ValueError("Missing required column(s): " + ", ".join(missing))

    out = body[[cols["Date"], cols["PID"], cols["First Name"], cols["Last Name"],
                cols["Center"], cols["General Service"], cols["Detailed Service"],
                cols["Service Type"], cols["Result"]]].copy()
    out.columns = list(cols.keys())

    # Date cleanup
    out["Date"] = pd.to_datetime(out["Date"], errors="coerce")
    out = out[out["Date"].notna()].reset_index(drop=True)

    # Filter out referrals before 8/1/2025
    cutoff = pd.Timestamp("2025-08-01")
    out = out[out["Date"] >= cutoff].reset_index(drop=True)

    # Format date as mm/dd/yy
    out["Date"] = out["Date"].dt.strftime("%m/%d/%y")

    return out

# ----------------------------
# Excel Writer (same styling)
# ----------------------------
def to_styled_excel(df: pd.DataFrame) -> bytes:
    def col_letter(n: int) -> str:
        s = ""
        while n >= 0:
            s = chr(n % 26 + 65) + s
            n = n // 26 - 1
        return s

    output = io.BytesIO()
    with pd.ExcelWriter(output, engine="xlsxwriter") as writer:
        df.to_excel(writer, index=False, sheet_name="Head Start Services & Referrals", startrow=3)
        wb = writer.book
        ws = writer.sheets["Head Start Services & Referrals"]

        ws.hide_gridlines(0)
        ws.set_row(0, 24); ws.set_row(1, 22); ws.set_row(2, 20)

        # Logo
        logo = Path("header_logo.png")
        if logo.exists():
            ws.set_column(1, 1, 6)
            ws.insert_image(0, 1, str(logo), {"x_offset":2,"y_offset":2,"x_scale":0.53,"y_scale":0.53,"object_position":1})

        # Titles
        now_ct = datetime.now(ZoneInfo("America/Chicago"))
        date_str = now_ct.strftime("%m.%d.%y %I:%M %p CT")
        title_fmt = wb.add_format({"bold": True, "font_size": 14, "align": "center"})
        subtitle_fmt = wb.add_format({"bold": True, "font_size": 12, "align": "center"})
        red_fmt = wb.add_format({"bold": True, "font_size": 12, "font_color": "#C00000"})
        last_col_0 = len(df.columns)-1
        last_col_letter = col_letter(last_col_0)

        ws.merge_range(0,2,0,last_col_0,"Hidalgo County Head Start Program", title_fmt)
        ws.merge_range(1,2,1,last_col_0,"",subtitle_fmt)
        ws.write_rich_string(1,2,
            subtitle_fmt,"Head Start - 2025-2026 Services & Referrals as of ",
            red_fmt,f"({date_str})",subtitle_fmt)

        # Header row
        header_fmt = wb.add_format({
            "bold":True,"font_color":"white","bg_color":"#305496",
            "align":"center","valign":"vcenter","text_wrap":True,"border":1
        })
        ws.set_row(3,26)
        for c,col in enumerate(df.columns):
            ws.write(3,c,col,header_fmt)

        last_row_0 = len(df)+3
        last_excel_row = last_row_0+1
        ws.autofilter(3,0,last_row_0,last_col_0)

        widths = {
            "Date":12,"PID":14,"First Name":16,"Last Name":16,"Center":28,
            "General Service":28,"Detailed Service":32,"Service Type":18,"Result":18
        }
        for name,width in widths.items():
            if name in df.columns:
                idx = df.columns.get_loc(name)
                ws.set_column(idx,idx,width)

        border_all = wb.add_format({"border":1})
        ws.conditional_format(f"A4:{last_col_letter}{last_excel_row}",
            {"type":"formula","criteria":"TRUE","format":border_all})

        # Outer box
        top=wb.add_format({"top":2}); bottom=wb.add_format({"bottom":2})
        left=wb.add_format({"left":2}); right=wb.add_format({"right":2})
        ws.conditional_format(f"A1:{last_col_letter}1",{"type":"formula","criteria":"TRUE","format":top})
        ws.conditional_format(f"A1:A{last_excel_row}",{"type":"formula","criteria":"TRUE","format":left})
        ws.conditional_format(f"{last_col_letter}1:{last_col_letter}{last_excel_row}",{"type":"formula","criteria":"TRUE","format":right})
        ws.conditional_format(f"A{last_excel_row}:{last_col_letter}{last_excel_row}",{"type":"formula","criteria":"TRUE","format":bottom})
        ws.write(0,last_col_0,"",wb.add_format({"right":2,"top":2}))
        ws.write(1,last_col_0,"",wb.add_format({"right":2}))

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
