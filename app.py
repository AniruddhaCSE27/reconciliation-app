import streamlit as st
import pandas as pd
from collections import Counter
from io import BytesIO

# ==========================================
# PAGE CONFIG
# ==========================================
st.set_page_config(
    page_title="Bank Reconciliation Tool",
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
    <div class="hero-title">💰 Smart Bank Reconciliation Tool</div>
    <div class="hero-subtitle">
        Upload your Excel file, auto-detect headers and columns, compare transactions,
        and generate a professional mismatch report with row-level references.
    </div>
</div>
""", unsafe_allow_html=True)

# ==========================================
# SIDEBAR
# ==========================================
with st.sidebar:
    st.markdown("## ⚙️ How to Use")
    st.write("1. Upload the Excel file")
    st.write("2. Choose comparison mode")
    st.write("3. Review dashboard and mismatches")
    st.write("4. Download CSV or Excel report")

    st.markdown("---")
    st.markdown("## ✅ Supported Modes")
    st.write("- Deposit vs Debit")
    st.write("- Withdrawal vs Credit")

    st.markdown("---")
    st.markdown("## ✅ Smart Features")
    st.write("- Auto header detection")
    st.write("- Flexible column matching")
    st.write("- Row-level mismatch tracing")
    st.write("- CSV and Excel export")

# ==========================================
# HELPERS
# ==========================================
def clean_columns(df):
    df.columns = df.columns.astype(str).str.strip()
    df.columns = df.columns.str.replace("\n", " ", regex=False)
    df.columns = df.columns.str.replace(r"\s+", " ", regex=True)
    return df

def normalize_name(name):
    return (
        str(name)
        .strip()
        .lower()
        .replace("\n", " ")
        .replace(" ", "")
        .replace("_", "")
    )

def find_matching_column(columns, aliases):
    normalized_map = {normalize_name(col): col for col in columns}
    for alias in aliases:
        alias_norm = normalize_name(alias)
        if alias_norm in normalized_map:
            return normalized_map[alias_norm]

    # fallback: partial match
    for col in columns:
        col_norm = normalize_name(col)
        for alias in aliases:
            alias_norm = normalize_name(alias)
            if alias_norm in col_norm or col_norm in alias_norm:
                return col
    return None

def read_sheet_with_header_detection(file_bytes, sheet_name, aliases_to_find, max_header_rows=12):
    for header_row in range(max_header_rows):
        try:
            temp_df = pd.read_excel(BytesIO(file_bytes), sheet_name=sheet_name, header=header_row)
            temp_df = clean_columns(temp_df)

            found_any = False
            for alias_group in aliases_to_find:
                match = find_matching_column(temp_df.columns, alias_group)
                if match is not None:
                    found_any = True
                    break

            if found_any:
                return temp_df, header_row
        except Exception:
            continue

    return None, None

def clean_numeric(series):
    return pd.to_numeric(
        series.astype(str)
        .str.replace(",", "", regex=False)
        .str.replace("₹", "", regex=False)
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
# COLUMN ALIASES
# ==========================================
DEPOSIT_ALIASES = [
    "Deposit Amt (INR)",
    "Deposit Amount (INR)",
    "Deposit Amt",
    "Deposit Amount",
    "Deposit"
]

WITHDRAWAL_ALIASES = [
    "Withdrawal Amt (INR)",
    "Withdrawal Amount (INR)",
    "Withdrawal Amt",
    "Withdrawal Amount",
    "Withdrawal",
    "Withdrawals"
]

DEBIT_ALIASES = [
    "Debit (LC)",
    "Debit(LC)",
    "Debit LC",
    "Debit Amount",
    "Debit"
]

CREDIT_ALIASES = [
    "Credit (LC)",
    "Credit(LC)",
    "Credit LC",
    "Credit Amount",
    "Credit"
]

# ==========================================
# FILE UPLOAD + MODE
# ==========================================
st.markdown('<div class="info-card"><div class="section-title">📤 Upload File</div></div>', unsafe_allow_html=True)

uploaded_file = st.file_uploader("Upload Excel File", type=["xlsx"])

comparison_mode = st.selectbox(
    "Select Comparison Type",
    ["Deposit vs Debit", "Withdrawal vs Credit"]
)

if uploaded_file:
    try:
        file_bytes = uploaded_file.read()

        # Read both sheets with flexible header detection
        sheet1, sheet1_header = read_sheet_with_header_detection(
            file_bytes=file_bytes,
            sheet_name="Sheet1",
            aliases_to_find=[DEPOSIT_ALIASES, WITHDRAWAL_ALIASES],
            max_header_rows=25
        )

        sheet2, sheet2_header = read_sheet_with_header_detection(
            file_bytes=file_bytes,
            sheet_name="Sheet2",
            aliases_to_find=[DEBIT_ALIASES, CREDIT_ALIASES],
            max_header_rows=12
        )

        if sheet1 is None:
            st.error("❌ Could not detect the correct header row in Sheet1.")
            st.stop()

        if sheet2 is None:
            st.error("❌ Could not detect the correct header row in Sheet2.")
            st.stop()

        # Choose columns by mode
        if comparison_mode == "Deposit vs Debit":
            left_label = "Deposit"
            right_label = "Debit"
            left_aliases = DEPOSIT_ALIASES
            right_aliases = DEBIT_ALIASES
        else:
            left_label = "Withdrawal"
            right_label = "Credit"
            left_aliases = WITHDRAWAL_ALIASES
            right_aliases = CREDIT_ALIASES

        left_col = find_matching_column(sheet1.columns, left_aliases)
        right_col = find_matching_column(sheet2.columns, right_aliases)

        if left_col is None:
            st.error(f"❌ {left_label} column not found in Sheet1.")
            st.write("Available Sheet1 columns:", list(sheet1.columns))
            st.stop()

        if right_col is None:
            st.error(f"❌ {right_label} column not found in Sheet2.")
            st.write("Available Sheet2 columns:", list(sheet2.columns))
            st.stop()

        # Clean numeric columns
        sheet1[left_col] = clean_numeric(sheet1[left_col])
        sheet2[right_col] = clean_numeric(sheet2[right_col])

        # Keep non-zero rows and preserve original row refs
        sheet1_filtered = sheet1[sheet1[left_col] != 0].reset_index()
        sheet2_filtered = sheet2[sheet2[right_col] != 0].reset_index()

        # Reconciliation
        left_count = Counter(sheet1_filtered[left_col])
        right_count = Counter(sheet2_filtered[right_col])

        all_amounts = sorted(set(left_count.keys()) | set(right_count.keys()))
        result = []

        for amt in all_amounts:
            left_amt_count = left_count.get(amt, 0)
            right_amt_count = right_count.get(amt, 0)

            if left_amt_count != right_amt_count:
                left_rows = sheet1_filtered[sheet1_filtered[left_col] == amt]["index"].tolist()
                right_rows = sheet2_filtered[sheet2_filtered[right_col] == amt]["index"].tolist()

                if right_amt_count > left_amt_count:
                    missing_where = f"Missing in {left_col}"
                    missing_count = right_amt_count - left_amt_count
                else:
                    missing_where = f"Missing in {right_col}"
                    missing_count = left_amt_count - right_amt_count

                result.append({
                    "Amount": amt,
                    f"{left_label} Count": left_amt_count,
                    f"{right_label} Count": right_amt_count,
                    "Missing Where": missing_where,
                    "Missing Count": missing_count,
                    "Sheet1 Rows": left_rows,
                    "Sheet2 Rows": right_rows
                })

        result_df = pd.DataFrame(result)

        # Metrics
        total_left = len(sheet1_filtered)
        total_right = len(sheet2_filtered)
        mismatch_groups = len(result_df)
        total_missing_entries = int(result_df["Missing Count"].sum()) if not result_df.empty else 0
        status_text = "OK" if result_df.empty else "Check"

        st.success("File loaded and processed successfully ✔")
        st.subheader("📊 Dashboard")

        c1, c2, c3, c4, c5 = st.columns(5)
        c1.metric(left_label, total_left)
        c2.metric(right_label, total_right)
        c3.metric("Mismatch Groups", mismatch_groups)
        c4.metric("Total Missing Entries", total_missing_entries)
        c5.metric("Status", status_text)

        # Tabs
        tab1, tab2, tab3, tab4 = st.tabs(["Summary", "Mismatches", "Preview", "Debug"])

        with tab1:
            st.markdown("### Reconciliation Overview")

            a, b = st.columns([1.35, 1])

            with a:
                if not result_df.empty:
                    chart_df = result_df[["Amount", "Missing Count"]].copy()
                    chart_df["Amount"] = chart_df["Amount"].astype(str)
                    st.bar_chart(chart_df.set_index("Amount"))
                else:
                    st.success("All amounts matched ✔")

            with b:
                summary_df = pd.DataFrame({
                    "Metric": [
                        "Comparison Type",
                        "Detected Sheet1 Header Row",
                        "Detected Sheet2 Header Row",
                        "Detected Sheet1 Column",
                        "Detected Sheet2 Column",
                        "File Name",
                        "Mismatch Groups",
                        "Total Missing Entries"
                    ],
                    "Value": [
                        comparison_mode,
                        sheet1_header,
                        sheet2_header,
                        left_col,
                        right_col,
                        uploaded_file.name,
                        mismatch_groups,
                        total_missing_entries
                    ]
                })
                st.dataframe(summary_df, use_container_width=True, hide_index=True)

        with tab2:
            st.markdown("### Detailed Mismatch Report")

            if not result_df.empty:
                col1, col2 = st.columns(2)

                with col1:
                    filter_type = st.selectbox(
                        "Filter by Missing Where",
                        options=["All"] + sorted(result_df["Missing Where"].astype(str).unique().tolist())
                    )

                with col2:
                    sort_by = st.selectbox(
                        "Sort by",
                        options=["Amount", "Missing Count", f"{left_label} Count", f"{right_label} Count"]
                    )

                filtered_result_df = result_df.copy()

                if filter_type != "All":
                    filtered_result_df = filtered_result_df[
                        filtered_result_df["Missing Where"] == filter_type
                    ]

                filtered_result_df = filtered_result_df.sort_values(by=sort_by).reset_index(drop=True)

                st.dataframe(filtered_result_df, use_container_width=True)

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
            st.write("Comparison mode:", comparison_mode)
            st.write("Detected Sheet1 header row:", sheet1_header)
            st.write("Detected Sheet2 header row:", sheet2_header)
            st.write("Detected Sheet1 column:", left_col)
            st.write("Detected Sheet2 column:", right_col)
            st.write("Sheet1 columns:", list(sheet1.columns))
            st.write("Sheet2 columns:", list(sheet2.columns))

        # Footer
        st.markdown("""
        <div class="footer-box">
            🚀 Built by <b>Aniruddha Pathak</b> • Smart Reconciliation Tool
        </div>
        """, unsafe_allow_html=True)

    except Exception as e:
        st.error(f"Error: {e}")
