# services_referrals_tool.py
import io
from pathlib import Path
from datetime import datetime, date

import pandas as pd
import streamlit as st
import xlsxwriter
from xlsxwriter.utility import xl_col_to_name

# -------------------------
# Streamlit Page Settings
# -------------------------
st.set_page_config(page_title="HCHSP Services & Referrals Tool", layout="wide")

# -------------------------
# Helper Functions
# -------------------------
def normalize_pid(val):
    try:
        return str(int(float(val)))
    except Exception:
        return str(val)

def extract_pir_code(detail_service):
    if pd.isna(detail_service):
        return None
    text = str(detail_service).upper()
    if "C.44" in text:
        start = text.find("C.44")
        return text[start:start + 6].strip()
    return None

def process_data(df, cutoff_date):
    df.columns = [str(c).strip() for c in df.columns]

    col_family = next((c for c in df.columns if "family" in c.lower()), None)
    col_pid    = next((c for c in df.columns if "pid" in c.lower()), None)
    col_first  = next((c for c in df.columns if "first" in c.lower()), None)
    col_last   = next((c for c in df.columns if "last" in c.lower()), None)
    col_center = next((c for c in df.columns if "center" in c.lower()), None)
    col_date   = next((c for c in df.columns if "date" in c.lower()), None)
    col_general= next((c for c in df.columns if "general" in c.lower()), None)
    col_detail = next((c for c in df.columns if "detail" in c.lower()), None)
    col_author = next((c for c in df.columns if "author" in c.lower()), None)
    col_result = next((c for c in df.columns if "result" in c.lower()), None)

    # Parse and filter dates
    df[col_date] = pd.to_datetime(df[col_date], errors="coerce")
    df = df[df[col_date] >= pd.to_datetime(cutoff_date)]

    # PIR code & rules
    df["PIR Code"] = df[col_detail].apply(extract_pir_code)

    valid_results = {"SERVICE COMPLETED", "SERVICE ONGOING"}
    df["Counts for PIR"] = "No"
    df["Reason (if not counted)"] = ""

    seen = set()
    for idx, row in df.iterrows():
        pid = normalize_pid(row[col_pid])
        pir_code = row["PIR Code"]
        result = (str(row[col_result]).upper().strip()
                  if pd.notna(row[col_result]) else "")

        if pir_code and result in valid_results:
            key = (pid, pir_code)
            if key not in seen:
                df.at[idx, "Counts for PIR"] = "Yes"
                seen.add(key)
            else:
                df.at[idx, "Reason (if not counted)"] = "Duplicate Entry"
        else:
            if not pir_code:
                df.at[idx, "Reason (if not counted)"] = "Missing Detailed Service"
            elif result not in valid_results:
                df.at[idx, "Reason (if not counted)"] = "Invalid Result"

    # Normalize PID display
    df[col_pid] = df[col_pid].apply(normalize_pid)

    return df, {
        "family": col_family,
        "pid": col_pid,
        "first": col_first,
        "last": col_last,
        "center": col_center,
        "date": col_date,
        "general": col_general,
        "detail": col_detail,
        "author": col_author,
        "result": col_result
    }

def build_summary(df, cols):
    pir_df = df[df["Counts for PIR"] == "Yes"].copy()
    if pir_df.empty:
        return pd.DataFrame(columns=[cols["general"], cols["detail"], "Distinct_Children", "PIR_Distinct_Families"])

    summary = (
        pir_df.groupby([cols["general"], cols["detail"]])
        .agg(
            Distinct_Children=(cols["pid"], "nunique"),
            PIR_Distinct_Families=(cols["family"], "nunique"),
        )
        .reset_index()
    )

    detailed_total = pd.DataFrame([{
        cols["general"]: "Detailed Services Total",
        cols["detail"]: "",
        "Distinct_Children": summary["Distinct_Children"].sum(),
        "PIR_Distinct_Families": summary["PIR_Distinct_Families"].sum(),
    }])

    summary = pd.concat([summary, detailed_total], ignore_index=True)
    return summary

def build_author_fix_list(df, cols):
    fix_df = df[df["Counts for PIR"] == "No"].copy()
    fix_df = fix_df[fix_df["Reason (if not counted)"].isin(
        ["Missing Detailed Service", "Invalid Result"]
    )]
    if fix_df.empty:
        return pd.DataFrame(columns=["Author", "Reason", "PIDs to Fix"])

    grouped = (
        fix_df.groupby([cols["author"], "Reason (if not counted)"])[cols["pid"]]
        .apply(lambda x: ", ".join(sorted(set(map(str, x)))))
        .reset_index()
        .rename(columns={
            cols["author"]: "Author",
            "Reason (if not counted)": "Reason",
            cols["pid"]: "PIDs to Fix"
        })
    )
    return grouped

def write_excel(df, summary, author_fix, cols, cutoff_date):
    output = io.BytesIO()
    with pd.ExcelWriter(output, engine="xlsxwriter") as writer:
        # Sheet names
        sheet1 = "Services & Referrals"
        sheet2 = "PIR Summary"
        sheet3 = "Author Fix List"

        # Sheet 1
        df_out = df[[cols["family"], cols["pid"], cols["last"], cols["first"],
                     cols["center"], cols["date"], cols["general"], cols["detail"],
                     cols["author"], cols["result"], "Counts for PIR", "Reason (if not counted)"]].copy()
        df_out.to_excel(writer, sheet_name=sheet1, index=False)

        # Sheet 2
        summary.to_excel(writer, sheet_name=sheet2, index=False)

        # Sheet 3
        author_fix.to_excel(writer, sheet_name=sheet3, index=False)

        wb = writer.book
        fmt_header = wb.add_format({"bold": True, "bg_color": "#1F4E78", "color": "white"})
        fmt_total = wb.add_format({"bold": True, "bg_color": "#16365C", "color": "white"})
        fmt_date = wb.add_format({"num_format": "mm/dd/yyyy"})

        # Format each sheet separately
        ws1 = writer.sheets[sheet1]
        ws1.freeze_panes(1, 0)
        ws1.set_row(0, None, fmt_header)
        ws1.autofilter(0, 0, 0, df_out.shape[1] - 1)

        # Date column format (if present)
        try:
            date_col_idx = list(df_out.columns).index(cols["date"])
            ws1.set_column(date_col_idx, date_col_idx, 12, fmt_date)
        except ValueError:
            pass

        # Total row (count visible rows in a key column)
        last_row_1 = len(df_out) + 1  # 1-based row in Excel (header is row 1)
        key_col_idx = list(df_out.columns).index(cols["general"])
        key_col_letter = xl_col_to_name(key_col_idx)
        ws1.write(last_row_1, 0, "Total", fmt_total)
        ws1.write_formula(
            last_row_1, key_col_idx,
            f'=SUBTOTAL(103,{sheet1}!{key_col_letter}2:{key_col_letter}{last_row_1})',
            fmt_total
        )

        # Sheet 2 formatting
        ws2 = writer.sheets[sheet2]
        if not summary.empty:
            ws2.freeze_panes(1, 0)
            ws2.set_row(0, None, fmt_header)
            ws2.autofilter(0, 0, 0, summary.shape[1] - 1)

        # Sheet 3 formatting
        ws3 = writer.sheets[sheet3]
        if not author_fix.empty:
            ws3.freeze_panes(1, 0)
            ws3.set_row(0, None, fmt_header)
            ws3.autofilter(0, 0, 0, author_fix.shape[1] - 1)

    return output.getvalue()

# ----------------------------
# Header (Streamlit UI)
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
        "<p style='text-align:center; font-size:16px; margin-top:0;'>Upload your Services & Referrals export, choose a cutoff date, and download the formatted report.</p>",
        unsafe_allow_html=True,
    )

# ----------------------------
# Controls
# ----------------------------
with st.sidebar:
    st.subheader("Settings")
    cutoff = st.date_input(
        "Cutoff date (include records on/after):",
        value=date(2025, 8, 15)  # adjust default if needed
    )

uploaded_file = st.file_uploader(
    "Upload Excel file (.xlsx) with Services & Referrals data",
    type=["xlsx"]
)

# ----------------------------
# Main
# ----------------------------
if uploaded_file is None:
    st.info("Upload your Excel file to begin.")
else:
    try:
        df_in = pd.read_excel(uploaded_file)
        processed, cols = process_data(df_in, pd.to_datetime(cutoff))
        summary = build_summary(processed, cols)
        author_fix = build_author_fix_list(processed, cols)

        st.subheader("Preview: Services & Referrals (first 20 rows)")
        st.dataframe(processed.head(20), use_container_width=True)

        st.subheader("Preview: PIR Summary")
        st.dataframe(summary, use_container_width=True)

        st.subheader("Preview: Author Fix List")
        st.dataframe(author_fix, use_container_width=True)

        excel_bytes = write_excel(processed, summary, author_fix, cols, pd.to_datetime(cutoff))
        st.download_button(
            label="📥 Download Excel Report",
            data=excel_bytes,
            file_name="HCHSP_Services_Referrals_Report.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )
    except Exception as e:
        st.error(f"There was an error processing your file: {e}")

    )

