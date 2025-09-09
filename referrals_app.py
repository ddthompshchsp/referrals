# services_referrals_tool.py

import io
from pathlib import Path
from datetime import datetime
import pandas as pd
import streamlit as st
import xlsxwriter

# -------------------------
# Streamlit Page Settings
# -------------------------
st.set_page_config(page_title="HCHSP Services & Referrals Tool", layout="wide")

# -------------------------
# Helper Functions
# -------------------------
def normalize_pid(val):
    """Ensure PIDs display as integers (no decimals)."""
    try:
        return str(int(float(val)))
    except:
        return str(val)

def extract_pir_code(detail_service):
    """Extract PIR code (e.g., C.44N) from Detailed Service text."""
    if pd.isna(detail_service):
        return None
    text = str(detail_service).upper()
    if "C.44" in text:
        start = text.find("C.44")
        return text[start:start+6].strip()
    return None

def process_data(df, cutoff_date):
    # Normalize column names
    df.columns = [str(c).strip() for c in df.columns]

    # Identify needed columns (flexible matching)
    col_family = next((c for c in df.columns if "family" in c.lower()), None)
    col_pid = next((c for c in df.columns if "pid" in c.lower()), None)
    col_first = next((c for c in df.columns if "first" in c.lower()), None)
    col_last = next((c for c in df.columns if "last" in c.lower()), None)
    col_center = next((c for c in df.columns if "center" in c.lower()), None)
    col_date = next((c for c in df.columns if "date" in c.lower()), None)
    col_general = next((c for c in df.columns if "general" in c.lower()), None)
    col_detail = next((c for c in df.columns if "detail" in c.lower()), None)
    col_author = next((c for c in df.columns if "author" in c.lower()), None)
    col_result = next((c for c in df.columns if "result" in c.lower()), None)

    # Parse dates
    df[col_date] = pd.to_datetime(df[col_date], errors="coerce")

    # Apply cutoff
    df = df[df[col_date] >= cutoff_date]

    # Add PIR Code
    df["PIR Code"] = df[col_detail].apply(extract_pir_code)

    # Apply strict PIR rules
    valid_results = ["SERVICE COMPLETED", "SERVICE ONGOING"]
    df["Counts for PIR"] = "No"
    df["Reason (if not counted)"] = ""

    seen = set()
    for idx, row in df.iterrows():
        pid = normalize_pid(row[col_pid])
        pir_code = row["PIR Code"]
        result = str(row[col_result]).upper().strip() if pd.notna(row[col_result]) else ""

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

    # Normalize PIDs
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
    # Filter PIR counts only
    pir_df = df[df["Counts for PIR"] == "Yes"]

    summary = (
        pir_df.groupby([cols["general"], cols["detail"]])
        .agg(
            Distinct_Children=(cols["pid"], "nunique"),
            PIR_Distinct_Families=(cols["family"], "nunique"),
        )
        .reset_index()
    )

    # Add totals
    detailed_total = pd.DataFrame([{
        cols["general"]: "Detailed Services Total",
        cols["detail"]: "",
        "Distinct_Children": summary["Distinct_Children"].sum(),
        "PIR_Distinct_Families": summary["PIR_Distinct_Families"].sum()
    }])

    summary = pd.concat([summary, detailed_total], ignore_index=True)

    return summary

def build_author_fix_list(df, cols):
    fix_df = df[df["Counts for PIR"] == "No"].copy()
    fix_df = fix_df[fix_df["Reason (if not counted)"].isin(
        ["Missing Detailed Service", "Invalid Result"]
    )]

    grouped = fix_df.groupby([cols["author"], "Reason (if not counted)"])[cols["pid"]].apply(
        lambda x: ", ".join(sorted(set(x)))
    ).reset_index()

    grouped = grouped.rename(columns={
        cols["author"]: "Author",
        "Reason (if not counted)": "Reason",
        cols["pid"]: "PIDs to Fix"
    })

    return grouped

def write_excel(df, summary, author_fix, cols, cutoff_date):
    output = io.BytesIO()
    with pd.ExcelWriter(output, engine="xlsxwriter") as writer:
        # Services & Referrals sheet
        sheet1 = "Services & Referrals"
        df_out = df[[cols["family"], cols["pid"], cols["last"], cols["first"],
                     cols["center"], cols["date"], cols["general"], cols["detail"],
                     cols["author"], cols["result"], "Counts for PIR", "Reason (if not counted)"]]
        df_out.to_excel(writer, sheet_name=sheet1, index=False)

        # PIR Summary sheet
        sheet2 = "PIR Summary"
        summary.to_excel(writer, sheet_name=sheet2, index=False)

        # Author Fix List sheet
        sheet3 = "Author Fix List"
        author_fix.to_excel(writer, sheet_name=sheet3, index=False)

        # Add formatting
        wb = writer.book
        fmt_header = wb.add_format({"bold": True, "bg_color": "#1F4E78", "color": "white"})
        fmt_total = wb.add_format({"bold": True, "bg_color": "#16365C", "color": "white"})
        fmt_date = wb.add_format({"num_format": "mm/dd/yyyy"})

        for sheet in [sheet1, sheet2, sheet3]:
            ws = writer.sheets[sheet]
            ws.freeze_panes(1, 0)
            ws.autofilter(0, 0, 0, df_out.shape[1] - 1)
            ws.set_row(0, None, fmt_header)

        # Format dates
        ws1 = writer.sheets[sheet1]
        date_col = list(df_out.columns).index(cols["date"])
        ws1.set_column(date_col, date_col, 12, fmt_date)

        # Add total row in Services
        last_row = len(df_out) + 1
        gen_col = list(df_out.columns).index(cols["general"])
        ws1.write(last_row, 0, "Total", fmt_total)
        ws1.write_formula(last_row, gen_col, f"=SUBTOTAL(103,{sheet1}!{chr(65+gen_col)}2:{chr(65+gen_col)}{last_row})", fmt_total)

    return output.getvalue()

# -------------------------
# Streamlit UI
# -------------------------

st.set_page_config(page_title="HCHSP Enrollment", layout="wide")

# ----------------------------
# Header (Streamlit UI only)
# ----------------------------
logo_path = Path("header_logo.png")
hdr_l, hdr_c, hdr_r = st.columns([1, 2, 1])
with hdr_c:
    if logo_path.exists():
        st.image(str(logo_path), width=320)
    st.markdown(
        "<h1 style='text-align:center; margin: 8px 0 4px;'>Hidalgo County Head Start — Enrollment Formatter</h1>",
        unsafe_allow_html=True,
    )
    st.markdown(
        """
        <p style='text-align:center; font-size:16px; margin-top:0;'>
        Upload the VF Average Funded Enrollment report and the 25–26 Applied/Accepted report.
        </p>
        """,
        unsafe_allow_html=True,
    )



if uploaded_file:
    df = pd.read_excel(uploaded_file)
    processed, cols = process_data(df, pd.to_datetime(cutoff_date))
    summary = build_summary(processed, cols)
    author_fix = build_author_fix_list(processed, cols)

    st.subheader("Preview: Services & Referrals")
    st.dataframe(processed.head(20))

    # Download button
    excel_bytes = write_excel(processed, summary, author_fix, cols, pd.to_datetime(cutoff_date))
    st.download_button(
        label="📥 Download Excel Report",
        data=excel_bytes,
        file_name="HCHSP_Services_Referrals_Report.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )

