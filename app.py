import streamlit as st
import pandas as pd
from collections import Counter
from io import BytesIO

# ==========================================
# PAGE CONFIG
# ==========================================
st.set_page_config(
    page_title="Deposit vs Debit Reconciliation Tool",
    page_icon="💰",
    layout="wide"
)

# ==========================================
# CUSTOM CSS
# ==========================================
st.markdown("""
<style>
    .main {
        background: linear-gradient(180deg, #f8fafc 0%, #eef4ff 100%);
    }

    .block-container {
        padding-top: 1.5rem;
        padding-bottom: 2rem;
        max-width: 1400px;
    }

    .hero-card {
        background: linear-gradient(135deg, #0f172a, #1d4ed8);
        padding: 28px 32px;
        border-radius: 20px;
        color: white;
        box-shadow: 0 10px 25px rgba(0,0,0,0.15);
        margin-bottom: 1rem;
    }

    .hero-title {
        font-size: 2rem;
        font-weight: 700;
        margin-bottom: 0.35rem;
    }

    .hero-subtitle {
        font-size: 1rem;
        opacity: 0.92;
    }

    .info-card {
        background: white;
        padding: 18px 20px;
        border-radius: 16px;
        box-shadow: 0 6px 16px rgba(15, 23, 42, 0.08);
        border: 1px solid #e2e8f0;
        margin-bottom: 0.75rem;
    }

    .section-title {
        font-size: 1.1rem;
        font-weight: 700;
        color: #0f172a;
        margin-bottom: 0.5rem;
    }

    .footer-box {
        margin-top: 25px;
        padding: 16px 20px;
        border-radius: 16px;
        background: white;
        border: 1px solid #e2e8f0;
        color: #334155;
        text-align: center;
        font-weight: 500;
        box-shadow: 0 6px 16px rgba(15, 23, 42, 0.05);
    }

    div[data-testid="stMetric"] {
        background: white;
        border: 1px solid #e2e8f0;
        padding: 16px;
        border-radius: 18px;
        box-shadow: 0 8px 18px rgba(15, 23, 42, 0.07);
    }

    div[data-testid="stDataFrame"] {
        border-radius: 14px;
        overflow: hidden;
    }

    .stDownloadButton > button,
    .stButton > button {
        border-radius: 12px;
        font-weight: 600;
        padding: 0.5rem 1rem;
    }
</style>
""", unsafe_allow_html=True)

# ==========================================
# HEADER
# ==========================================
st.markdown("""
<div class="hero-card">
    <div class="hero-title">💰 Deposit vs Debit Reconciliation Tool</div>
    <div class="hero-subtitle">
        Upload your bank Excel file, auto-detect format, compare Deposit vs Debit entries,
        and generate a professional mismatch report instantly.
    </div>
</div>
""", unsafe_allow_html=True)

# ==========================================
# SIDEBAR
# ==========================================
with st.sidebar:
    st.markdown("## ⚙️ How to Use")
    st.write("1. Upload the Excel file")
    st.write("2. Tool auto-detects format")
    st.write("3. Review summary, mismatches, and row references")
    st.write("4. Download report")

    st.markdown("---")
    st.markdown("## ✅ Features")
    st.write("- Auto header detection")
    st.write("- Smart column detection")
    st.write("- Dashboard metrics")
    st.write("- Row-level mismatch references")
    st.write("- CSV & Excel export")

    st.markdown("---")
    st.info("Tip: 'Mismatch Groups' means how many different amount values have mismatches. 'Total Missing Entries' means the actual unmatched row count.")

# ==========================================
# HELPERS
# ==========================================
def clean_columns(df):
    df.columns = df.columns.astype(str).str.strip()
    df.columns = df.columns.str.replace("\n", " ", regex=False)
    df.columns = df.columns.str.replace(r"\s+", " ", regex=True)
    return df

def read_sheet2_correctly(file_bytes):
    """
    Try different header rows for Sheet2 and return the first one
    where a debit-like column is found.
    """
    for i in range(6):
        try:
            temp = pd.read_excel(BytesIO(file_bytes), sheet_name="Sheet2", header=i)
            temp = clean_columns(temp)
            cols = [str(c).lower() for c in temp.columns]

            if any(("debit" in c) or ("withdraw" in c) for c in cols):
                return temp, i
        except Exception:
            continue

    return None, None

def find_deposit_col(cols):
    possible_cols = [
        "Deposit Amt (INR)",
        "Deposit Amount (INR)",
        "Deposit Amt",
        "Deposit Amount",
        "Deposit"
    ]
    for c in possible_cols:
        if c in cols:
            return c
    return None

def find_debit_col(cols):
    possible_cols = [
        "Debit (LC)",
        "Debit(LC)",
        "Debit",
        "Debit LC",
        "Debit Amount",
        "Withdrawal Amount",
        "Withdrawals"
    ]
    for c in possible_cols:
        if c in cols:
            return c
    return None

def clean_numeric(series):
    return pd.to_numeric(
        series.astype(str)
        .str.replace(",", "", regex=False)
        .str.strip(),
        errors="coerce"
    ).fillna(0)

def to_excel(df):
    output = BytesIO()
    with pd.ExcelWriter(output, engine="openpyxl") as writer:
        df.to_excel(writer, index=False, sheet_name="Reconciliation_Result")
    output.seek(0)
    return output.getvalue()

# ==========================================
# FILE UPLOAD
# ==========================================
st.markdown('<div class="info-card"><div class="section-title">📤 Upload File</div></div>', unsafe_allow_html=True)

uploaded_file = st.file_uploader("Upload Excel File", type=["xlsx"])

if uploaded_file:
    try:
        file_bytes = uploaded_file.read()

        # ==========================================
        # READ SHEETS
        # ==========================================
        sheet1 = pd.read_excel(BytesIO(file_bytes), sheet_name="Sheet1", header=16)
        sheet1 = clean_columns(sheet1)

        sheet2, header_row = read_sheet2_correctly(file_bytes)

        if sheet2 is None:
            st.error("❌ Could not detect the correct header row in Sheet2.")
            st.stop()

        # ==========================================
        # DETECT REQUIRED COLUMNS
        # ==========================================
        deposit_col = find_deposit_col(sheet1.columns)
        debit_col = find_debit_col(sheet2.columns)

        if deposit_col is None:
            st.error("❌ Deposit column not found in Sheet1.")
            st.write("Available Sheet1 columns:", list(sheet1.columns))
            st.stop()

        if debit_col is None:
            st.error("❌ Debit column not found in Sheet2.")
            st.write("Available Sheet2 columns:", list(sheet2.columns))
            st.stop()

        # ==========================================
        # CLEAN DATA
        # ==========================================
        sheet1[deposit_col] = clean_numeric(sheet1[deposit_col])
        sheet2[debit_col] = clean_numeric(sheet2[debit_col])

        # Keep only non-zero values and preserve original row positions
        sheet1_filtered = sheet1[sheet1[deposit_col] != 0].reset_index()
        sheet2_filtered = sheet2[sheet2[debit_col] != 0].reset_index()

        # ==========================================
        # RECONCILIATION LOGIC
        # ==========================================
        deposit_count = Counter(sheet1_filtered[deposit_col])
        debit_count = Counter(sheet2_filtered[debit_col])

        all_amt = sorted(set(deposit_count.keys()) | set(debit_count.keys()))

        result = []

        for amt in all_amt:
            dep_count = deposit_count.get(amt, 0)
            deb_count = debit_count.get(amt, 0)

            if dep_count != deb_count:
                deposit_rows = sheet1_filtered[
                    sheet1_filtered[deposit_col] == amt
                ]["index"].tolist()

                debit_rows = sheet2_filtered[
                    sheet2_filtered[debit_col] == amt
                ]["index"].tolist()

                if deb_count > dep_count:
                    missing_where = f"Missing in {deposit_col}"
                    missing_count = deb_count - dep_count
                else:
                    missing_where = f"Missing in {debit_col}"
                    missing_count = dep_count - deb_count

                result.append({
                    "Amount": amt,
                    "Deposit Count": dep_count,
                    "Debit Count": deb_count,
                    "Missing Where": missing_where,
                    "Missing Count": missing_count,
                    "Sheet1 Rows": deposit_rows,
                    "Sheet2 Rows": debit_rows
                })

        result_df = pd.DataFrame(result)

        # ==========================================
        # KPI METRICS
        # ==========================================
        st.success("File loaded and processed successfully ✔")
        st.subheader("📊 Dashboard")

        total_deposits = len(sheet1_filtered)
        total_debits = len(sheet2_filtered)
        mismatch_groups = len(result_df)
        total_missing_entries = int(result_df["Missing Count"].sum()) if not result_df.empty else 0
        status_text = "OK" if result_df.empty else "Check"

        c1, c2, c3, c4, c5 = st.columns(5)
        c1.metric("Deposits", total_deposits)
        c2.metric("Debits", total_debits)
        c3.metric("Mismatch Groups", mismatch_groups)
        c4.metric("Total Missing Entries", total_missing_entries)
        c5.metric("Status", status_text)

        # ==========================================
        # TABS
        # ==========================================
        tab1, tab2, tab3, tab4 = st.tabs(["Summary", "Mismatches", "Preview", "Debug"])

        with tab1:
            st.markdown("### Reconciliation Overview")

            left, right = st.columns([1.35, 1])

            with left:
                if not result_df.empty:
                    chart_df = result_df[["Amount", "Missing Count"]].copy()
                    chart_df["Amount"] = chart_df["Amount"].astype(str)
                    st.bar_chart(chart_df.set_index("Amount"))
                else:
                    st.success("All amounts matched ✔")

            with right:
                summary_df = pd.DataFrame({
                    "Metric": [
                        "Detected Deposit Column",
                        "Detected Debit Column",
                        "Detected Sheet2 Header Row",
                        "File Name",
                        "Mismatch Groups",
                        "Total Missing Entries"
                    ],
                    "Value": [
                        deposit_col,
                        debit_col,
                        header_row,
                        uploaded_file.name,
                        mismatch_groups,
                        total_missing_entries
                    ]
                })
                st.dataframe(summary_df, use_container_width=True, hide_index=True)

            if not result_df.empty:
                st.info(
                    f"There are {mismatch_groups} mismatch groups across different amount values, "
                    f"with a total of {total_missing_entries} missing/unmatched entries."
                )

        with tab2:
            st.markdown("### Detailed Mismatch Report")

            if not result_df.empty:
                col_a, col_b = st.columns(2)

                with col_a:
                    filter_type = st.selectbox(
                        "Filter by Missing Where",
                        options=["All"] + sorted(result_df["Missing Where"].astype(str).unique().tolist())
                    )

                with col_b:
                    sort_by = st.selectbox(
                        "Sort by",
                        options=["Amount", "Missing Count", "Deposit Count", "Debit Count"]
                    )

                filtered_result_df = result_df.copy()

                if filter_type != "All":
                    filtered_result_df = filtered_result_df[
                        filtered_result_df["Missing Where"] == filter_type
                    ]

                filtered_result_df = filtered_result_df.sort_values(by=sort_by).reset_index(drop=True)

                st.dataframe(filtered_result_df, use_container_width=True)

                st.markdown("#### Download Report")
                d1, d2 = st.columns(2)

                with d1:
                    st.download_button(
                        "⬇ Download CSV",
                        filtered_result_df.to_csv(index=False).encode("utf-8"),
                        "reconciliation_result.csv",
                        "text/csv"
                    )

                with d2:
                    st.download_button(
                        "⬇ Download Excel",
                        to_excel(filtered_result_df),
                        "reconciliation_result.xlsx",
                        "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                    )

                st.markdown("#### Understanding Row References")
                st.caption(
                    "Sheet1 Rows and Sheet2 Rows show the original row positions where that amount appears "
                    "after loading the file. This helps trace exactly where mismatches belong."
                )

            else:
                st.success("All amounts matched ✔ No mismatches found.")

        with tab3:
            st.markdown("### Raw Data Preview")

            p1, p2 = st.columns(2)

            with p1:
                st.markdown("**Sheet1 Preview**")
                st.dataframe(sheet1_filtered.head(20), use_container_width=True)

            with p2:
                st.markdown("**Sheet2 Preview**")
                st.dataframe(sheet2_filtered.head(20), use_container_width=True)

        with tab4:
            st.markdown("### Debug Information")
            st.write("Detected Sheet2 header row:", header_row)
            st.write("Detected deposit column:", deposit_col)
            st.write("Detected debit column:", debit_col)
            st.write("Sheet1 columns:", list(sheet1.columns))
            st.write("Sheet2 columns:", list(sheet2.columns))

        # ==========================================
        # FOOTER
        # ==========================================
        st.markdown("""
        <div class="footer-box">
            🚀 Built by <b>Aniruddha Pathak</b> • Smart Reconciliation Tool
        </div>
        """, unsafe_allow_html=True)

    except Exception as e:
        st.error(f"Error: {e}")
