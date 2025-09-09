
import io
import re
from pathlib import Path
from datetime import datetime
from zoneinfo import ZoneInfo

import pandas as pd
import streamlit as st

st.set_page_config(page_title="HCHSP Services & Referrals — Tool", layout="wide")

# =========================
# Header
# =========================
logo_path = Path("header_logo.png")  # place your logo file next to this script
hdr_l, hdr_c, hdr_r = st.columns([1, 2, 1])
with hdr_c:
    if logo_path.exists():
        st.image(str(logo_path), width=320)
    st.markdown("""
<h1 style='text-align:center; margin: 8px 0 4px;'>Hidalgo County Head Start — Services & Referrals</h1>
<p style='text-align:center; font-size:16px; margin-top:0;'>
Upload the <strong>10433</strong> Services/Referrals report (Excel). The tool will build:
<em>Services & Referrals</em>, <em>PIR Summary</em>, and <em>Author Fix List</em> with the exact formatting.
</p>
""", unsafe_allow_html=True)
st.divider()

# =========================
# Controls
# =========================
with st.sidebar:
    st.header("Settings")
    cutoff = st.date_input("Cutoff (Service Date on/after)", value=pd.to_datetime("2025-08-11")).strftime("%Y-%m-%d")
    st.caption("Only services with Service Date >= this date are included.")
    st.checkbox("Require 'PIR' in Detailed Service", value=True, key="require_pir")
    st.caption("PIR = must contain 'PIR' and a C.44 letter code.")

# =========================
# Upload
# =========================
inp_l, inp_c, inp_r = st.columns([1, 2, 1])
with inp_c:
    sref_file = st.file_uploader("Upload *10433.xlsx*", type=["xlsx"], key="sref")
    process = st.button("Process & Download")

# =========================
# Helpers
# =========================
def _clean_header(h: str) -> str:
    return re.sub(r"^(ST:|FD:)\s*", "", str(h).strip(), flags=re.I)

def _parse_to_dt(series: pd.Series) -> pd.Series:
    dt1 = pd.to_datetime(series, errors="coerce", infer_datetime_format=True)
    num = pd.to_numeric(series, errors="coerce")
    serial_mask = num.notna() & num.between(10000, 70000)
    dt2 = pd.Series(pd.NaT, index=series.index, dtype="datetime64[ns]")
    if serial_mask.any():
        dt2.loc[serial_mask] = pd.to_datetime(num.loc[serial_mask], unit="D", origin="1899-12-30", errors="coerce")
    dt = dt1.copy()
    dt[dt.isna()] = dt2[dt.isna()]
    return dt

def _extract_pir_code(text: str) -> str | None:
    if not isinstance(text, str): text = str(text)
    m = re.search(r'(?i)\bC\s*\.?\s*44\s*([a-z])\b', text)
    if m: return f"C.44 {m.group(1).lower()}"
    return None

def _format_pid(val) -> str:
    if pd.isna(val): return ""
    s = str(val).strip()
    if re.fullmatch(r"-?\d+\.0", s):
        try: return str(int(float(s)))
        except: return s
    if re.fullmatch(r"-?\d+\.\d+", s):
        try:
            f = float(s)
            if abs(f - int(f)) < 1e-9: return str(int(f))
            return s
        except: return s
    if re.fullmatch(r"-?\d+", s): return s
    return s

def _col_letter(idx: int) -> str:
    s = ""
    n = idx
    while n >= 0:
        s = chr(n % 26 + 65) + s
        n = n // 26 - 1
    return s

# =========================
# Excel writer
# =========================
def build_workbook(df_raw: pd.DataFrame, cutoff: str, require_pir: bool = True) -> bytes:
    # Map columns (header row is row 5 in most exports; this app expects headers present already)
    df = df_raw.copy()
    df.columns = [str(c).strip() for c in df.columns]

    # Column guesses
    FID_COL   = "Family ID"
    PID_COL   = "ST: Participant PID"
    LNAME_COL = "ST: Participant Last Name"
    FNAME_COL = "ST: Participant First Name"
    GEN_COL   = "FD: Services - General Service"
    DET_COL   = "FD: Services - Detail Service"
    RES_COL   = "FD: Services - Result"
    DATE_COL  = "FD: Services - Date"
    AUTH_COL  = next((c for c in df.columns if ("author" in c.lower() and "service" in c.lower())), None)
    if AUTH_COL is None:
        AUTH_COL = next((c for c in df.columns if any(k in c.lower() for k in ["author","staff","worker"])), None)
    CENTER_COL = next((c for c in df.columns if "center" in c.lower() or "campus" in c.lower()), None)

    # Date filter
    df[DATE_COL] = _parse_to_dt(df[DATE_COL])
    df = df[df[DATE_COL].notna() & (df[DATE_COL] >= pd.Timestamp(cutoff))].copy()

    # PIR flags
    df["_Result_norm"] = df[RES_COL].astype(str).str.strip().str.lower()
    valid_result = df["_Result_norm"].isin({"service ongoing","service completed"})
    has_pir = df[DET_COL].astype(str).str.contains("pir", case=False, na=False) if require_pir else True
    df["_has_PIR"] = has_pir
    df["_PIR_CODE"] = df[DET_COL].astype(str).map(_extract_pir_code)
    df["PID_norm"] = pd.to_numeric(df[PID_COL], errors="coerce")

    count_candidate = df["_has_PIR"] & valid_result & df["_PIR_CODE"].notna()
    dup_mask = pd.Series(False, index=df.index)
    if count_candidate.any():
        sub = pd.DataFrame({
            "pid": df.loc[count_candidate, "PID_norm"],
            "code": df.loc[count_candidate, "_PIR_CODE"].astype(str).str.strip().str.lower()
        }, index=df.index[count_candidate])
        dup_mask.loc[count_candidate] = sub.duplicated(subset=["pid","code"], keep="first").values

    df["Counts for PIR"] = (count_candidate & ~dup_mask).map({True:"Yes", False:"No"})

    # Reason column (NO "Missing PIR code in Detailed Service")
    def reason_fn(row):
        if row["Counts for PIR"] == "Yes":
            return ""
        gen = str(row.get(GEN_COL, "")).strip()
        det = str(row.get(DET_COL, "")).strip()
        res = str(row.get(RES_COL, "")).strip().lower()
        if gen == "" or gen.lower() == "nan":
            return "Missing General Service"
        if det == "" or det.lower() == "nan":
            return "Missing Detailed Service"
        if res not in {"service ongoing","service completed"}:
            return "Invalid/Missing Result"
        if row.name in dup_mask.index and dup_mask.loc[row.name]:
            return "Duplicate Entry"
        if pd.isna(row[DATE_COL]):
            return "Missing Service Date"
        return ""

    df["Reason (if not counted)"] = df.apply(reason_fn, axis=1)

    # Services & Referrals details
    cols = [FID_COL, PID_COL, LNAME_COL, FNAME_COL]
    if CENTER_COL: cols.append(CENTER_COL)
    cols += [DATE_COL, GEN_COL, DET_COL]
    if AUTH_COL: cols.append(AUTH_COL)
    cols += [RES_COL, "Counts for PIR", "Reason (if not counted)"]
    details = df[cols].copy()
    rename_map = {c: _clean_header(c) for c in details.columns if c not in ["Counts for PIR","Reason (if not counted)"]}
    details.rename(columns=rename_map, inplace=True)
    date_out = _clean_header(DATE_COL)
    pid_out  = _clean_header(PID_COL)
    details[date_out] = _parse_to_dt(details[date_out]).dt.strftime("%m/%d/%y")
    details[pid_out] = details[pid_out].apply(_format_pid)
    details = details.fillna("")

    # PIR Summary
    pir_rows = df[df["Counts for PIR"] == "Yes"].copy()
    pir_rows["_pid_norm"] = pd.to_numeric(pir_rows[PID_COL], errors="coerce")
    per_child = (pir_rows.dropna(subset=["_pid_norm"]
                ).drop_duplicates(subset=["_pid_norm", GEN_COL, "_PIR_CODE"]
                ).groupby([GEN_COL, DET_COL]).size()
                .rename("Distinct Children (PID)").reset_index())
    pir_rows[FID_COL] = pir_rows[FID_COL].astype(str).str.strip()
    per_family = (pir_rows.drop_duplicates(subset=[FID_COL, GEN_COL, "_PIR_CODE"]
                 ).groupby([GEN_COL, DET_COL]).size()
                 .rename("PIR (Distinct Families)").reset_index())
    summary = per_child.merge(per_family, on=[GEN_COL, DET_COL], how="outer").fillna(0)
    summary.rename(columns={GEN_COL:"GENERAL service", DET_COL:"DETAILED services"}, inplace=True)
    summary = summary[["GENERAL service","DETAILED services","Distinct Children (PID)","PIR (Distinct Families)"]]

    # Author Fix List (actionable only; PID list clean integers; add count column)
    author_col_name = _clean_header(AUTH_COL) if AUTH_COL else None
    actionable = {"Missing General Service","Missing Detailed Service","Invalid/Missing Result","Missing Service Date"}
    fix_rows = details[(details["Counts for PIR"]=="No") & (details["Reason (if not counted)"].isin(actionable))].copy()

    if author_col_name and author_col_name in fix_rows.columns:
        pids_by_group = (fix_rows.groupby([author_col_name, "Reason (if not counted)"])[_clean_header(PID_COL)]
                         .apply(lambda s: ", ".join(sorted({_format_pid(x) for x in s if str(x).strip() != ""}))))
        author_fix = pids_by_group.reset_index().rename(columns={_clean_header(PID_COL): "PIDs to Fix"})
        author_fix["Count of PIDs"] = author_fix["PIDs to Fix"].apply(lambda x: 0 if x=="" else len([p for p in x.split(", ") if p]))
    else:
        author_fix = pd.DataFrame(columns=[author_col_name or "Author","Reason (if not counted)","PIDs to Fix","Count of PIDs"])

    # Write to Excel (styled)
    buf = io.BytesIO()
    with pd.ExcelWriter(buf, engine="xlsxwriter") as writer:
        wb = writer.book
        hdr_fmt = wb.add_format({"bold":True,"font_color":"white","bg_color":"#305496","align":"center","valign":"vcenter","text_wrap":True,"border":1})
        border_all = wb.add_format({"border":1})
        title_fmt = wb.add_format({"bold": True, "font_size": 14, "align": "center"})
        subtitle_fmt = wb.add_format({"bold": True, "font_size": 12, "align": "center"})
        red_fmt = wb.add_format({"bold": True, "font_size": 12, "font_color": "#C00000"})
        light_red = wb.add_format({"bg_color":"#F8D7DA"})
        bold_center = wb.add_format({"bold":True,"align":"center"})

        # Sheet 1: Services & Referrals
        sheet1 = "Services & Referrals"
        details.to_excel(writer, index=False, sheet_name=sheet1, startrow=3)
        ws1 = writer.sheets[sheet1]
        ws1.hide_gridlines(0); ws1.set_row(0,24); ws1.set_row(1,22); ws1.set_row(2,20)

        # Logo + centered title + CT timestamp
        last_col_0 = details.shape[1] - 1
        if logo_path.exists():
            ws1.set_column(0, 0, 16)
            ws1.insert_image(0, 0, str(logo_path), {"x_offset":2, "y_offset":2, "x_scale":0.53, "y_scale":0.53, "object_position": 1})
        ws1.merge_range(0, 1, 0, last_col_0, "Hidalgo County Head Start Program", title_fmt)
        now_ct = datetime.now(ZoneInfo("America/Chicago")).strftime("%m/%d/%y %I:%M %p CT")
        ws1.merge_range(1, 1, 1, last_col_0, "", subtitle_fmt)
        ws1.write_rich_string(1, 1, subtitle_fmt, "Services & Referrals — 2025-2026 as of ", red_fmt, f"({now_ct})", subtitle_fmt)

        # Headers + layout
        ws1.set_row(3,26)
        for c, col in enumerate(details.columns):
            ws1.write(3, c, col, hdr_fmt)
        ws1.freeze_panes(4, 0)
        last_row_0 = len(details) + 3
        ws1.autofilter(3, 0, last_row_0, last_col_0)
        ws1.conditional_format(3, 0, last_row_0, last_col_0, {"type":"no_errors","format":border_all})

        # Column widths + blanks highlight
        def _set_widths(ws, cols):
            for idx, name in enumerate(cols):
                n=name.lower(); w=20
                if "center" in n or "campus" in n: w=26
                if "general service" in n: w=30
                if "detail service" in n or "detailed service" in n: w=40
                if "date" in n: w=14
                if "result" in n: w=20
                if "author" in n: w=24
                if "pid" in n or "family id" in n: w=16
                if "last name" in n or "first name" in n: w=22
                if "counts for pir" in n: w=18
                if "reason" in n: w=34
                ws.set_column(idx, idx, w)

        _set_widths(ws1, details.columns)

        name_to_idx = {name: idx for idx, name in enumerate(details.columns)}
        for name in details.columns:
            if ("general service" in name.lower() or "detail service" in name.lower() or
                "author" in name.lower() or "center" in name.lower() or "campus" in name.lower()):
                ws1.conditional_format(4, name_to_idx[name], last_row_0, name_to_idx[name], {"type":"blanks","format": light_red})

        # Dynamic TOTAL: label in column A, value under General Service
        helper_idx = last_col_0 + 1
        ws1.write(3, helper_idx, "_helper_")
        for r in range(4, last_row_0 + 1):
            ws1.write_number(r, helper_idx, 1)
        ws1.set_column(helper_idx, helper_idx, None, None, {"hidden":1})
        totals_row = last_row_0 + 1
        ws1.write(totals_row, 0, "Total", wb.add_format({"bold":True,"align":"right"}))
        helper_col_letter = _col_letter(helper_idx)
        try:
            gs_idx = next(i for i, h in enumerate(list(details.columns)) if "general service" in h.lower())
        except StopIteration:
            gs_idx = 5
        ws1.write_formula(totals_row, gs_idx, f"=SUBTOTAL(109,{helper_col_letter}5:{helper_col_letter}{last_row_0+1})", bold_center)

        # Sheet 2: PIR Summary
        sheet2 = "PIR Summary"
        summary.to_excel(writer, index=False, sheet_name=sheet2, startrow=1)
        ws2 = writer.sheets[sheet2]
        ws2.hide_gridlines(0)
        ws2.set_row(0, 24)
        if logo_path.exists():
            ws2.set_column(0, 0, 16)
            ws2.insert_image(0, 0, str(logo_path), {"x_offset":2, "y_offset":2, "x_scale":0.53, "y_scale":0.53, "object_position": 1})
        ws2.write(0, 1, "PIR Summary", wb.add_format({"bold": True, "font_size": 14, "align": "left"}))
        ws2.set_row(1,26)
        for c, col in enumerate(summary.columns):
            ws2.write(1, c, col, hdr_fmt)
        last_row2 = len(summary) + 1
        last_col2 = len(summary.columns) - 1
        ws2.autofilter(1, 0, last_row2, last_col2)
        ws2.conditional_format(1, 0, last_row2, last_col2, {"type": "no_errors", "format": border_all})
        _set_widths(ws2, summary.columns)

        # Summary totals
        start_excel_row = 3
        end_excel_row = last_row2 + 1
        def _col_letter2(i):
            s=""; n=i
            while n>=0:
                s=chr(n%26+65)+s; n=n//26-1
            return s
        children_col = _col_letter2(2)  # C
        families_col = _col_letter2(3)  # D
        total_fmt = wb.add_format({"bold": True, "bg_color": "#E2EFDA", "border": 1})
        ws2.write(last_row2 + 2, 1, "Detailed Services Total", total_fmt)
        ws2.write_formula(last_row2 + 2, 2, f"=SUBTOTAL(109,{children_col}{start_excel_row}:{children_col}{end_excel_row})", total_fmt)
        ws2.write_formula(last_row2 + 2, 3, f"=SUBTOTAL(109,{families_col}{start_excel_row}:{families_col}{end_excel_row})", total_fmt)

        c44_sum_fmt = wb.add_format({"bold": True, "bg_color": "#D9E1F2", "border": 1})
        ws2.write(last_row2 + 3, 1, "C.44 – Sum of PIR Families (TOTAL)", c44_sum_fmt)
        ws2.write_formula(last_row2 + 3, 3, f"=SUM({families_col}{start_excel_row}:{families_col}{end_excel_row})", c44_sum_fmt)

        # Sheet 3: Author Fix List
        sheet3 = "Author Fix List"
        author_fix.to_excel(writer, index=False, sheet_name=sheet3, startrow=1)
        ws3 = writer.sheets[sheet3]
        ws3.hide_gridlines(0)
        ws3.set_row(0, 24)
        if logo_path.exists():
            ws3.set_column(0, 0, 16)
            ws3.insert_image(0, 0, str(logo_path), {"x_offset":2, "y_offset":2, "x_scale":0.53, "y_scale":0.53, "object_position": 1})
        ws3.write(0, 1, "Author Fix List (Actionable only)", wb.add_format({"bold": True, "font_size": 14, "align": "left"}))
        ws3.set_row(1,26)
        for c, col in enumerate(author_fix.columns):
            ws3.write(1, c, col, hdr_fmt)
        ws3.autofilter(1, 0, len(author_fix) + 1, len(author_fix.columns) - 1)
        ws3.conditional_format(1, 0, len(author_fix) + 1, len(author_fix.columns) - 1, {"type":"no_errors","format": border_all})
        # Widths
        for idx, name in enumerate(author_fix.columns):
            w = 22
            if "reason" in name.lower(): w = 30
            if "pids" in name.lower():   w = 50
            ws3.set_column(idx, idx, w)

    return buf.getvalue()

# =========================
# Main
# =========================
if process and sref_file:
    try:
        raw = pd.read_excel(sref_file, header=4)  # standard 10433 export header row
        xlsx = build_workbook(raw, cutoff, require_pir=st.session_state.get("require_pir", True))
        st.download_button(
            "Download Styled Workbook (Excel)",
            data=xlsx,
            file_name="HCHSP_Services_Referrals_PIR.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        )
        st.success("Workbook generated. You can filter and the totals will follow.")
    except Exception as e:
        st.error(f"Processing error: {e}")


