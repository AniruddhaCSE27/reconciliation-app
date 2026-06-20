import streamlit as st
import pandas as pd
from collections import Counter
from io import BytesIO

st.set_page_config(
    page_title="Smart Bank Reconciliation Tool",
    page_icon="💰",
    layout="wide"
)

st.markdown("""
<style>
.block-container { padding-top: 1.5rem; max-width: 1400px; }
.hero-card {
    background: linear-gradient(135deg, #0f172a, #1d4ed8);
    padding: 28px 32px;
    border-radius: 20px;
    color: white;
    margin-bottom: 1rem;
}
.hero-title { font-size: 2rem; font-weight: 700; }
.hero-subtitle { font-size: 1rem; opacity: 0.92; }
.info-card {
    background: white;
    padding: 18px 20px;
    border-radius: 16px;
    margin-bottom: 0.75rem;
}
.section-title { font-size: 1.1rem; font-weight: 700; color: #0f172a; }
.footer-box {
    margin-top: 25px;
    padding: 16px 20px;
    border-radius: 16px;
    background: white;
    text-align: center;
    font-weight: 500;
}
</style>
""", unsafe_allow_html=True)

st.markdown("""
<div class="hero-card">
    <div class="hero-title">💰 Smart Bank Reconciliation Tool</div>
    <div class="hero-subtitle">
        Upload your Excel file, auto-detect sheets, headers and columns,
        compare transactions, and generate a professional mismatch report.
    </div>
</div>
""", unsafe_allow_html=True)

with st.sidebar:
    st.markdown("## ⚙️ How to Use")
    st.write("1. Upload Excel file")
    st.write("2. Select bank sheet and ERP sheet")
    st.write("3. Choose comparison mode")
    st.write("4. Download mismatch report")

    st.markdown("---")
    st.markdown("## ✅ Supported Modes")
    st.write("- Deposit vs Debit")
    st.write("- Withdrawal vs Credit")


# =========================
# Helper Functions
# =========================

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
        .replace(".", "")
        .replace("/", "")
        .replace("-", "")
        .replace("(", "")
        .replace(")", "")
    )


def find_matching_column(columns, aliases):
    normalized_map = {normalize_name(col): col for col in columns}

    for alias in aliases:
        alias_norm = normalize_name(alias)
        if alias_norm in normalized_map:
            return normalized_map[alias_norm]

    for col in columns:
        col_norm = normalize_name(col)
        for alias in aliases:
            alias_norm = normalize_name(alias)
            if alias_norm in col_norm or col_norm in alias_norm:
                return col

    return None


def read_sheet_with_header_detection(file_bytes, sheet_name, alias_groups, max_header_rows=60):
    best_df = None
    best_header = None
    best_score = 0

    for header_row in range(max_header_rows):
        try:
            temp_df = pd.read_excel(
                BytesIO(file_bytes),
                sheet_name=sheet_name,
                header=header_row
            )

            temp_df = clean_columns(temp_df)

            score = 0
            for aliases in alias_groups:
                match = find_matching_column(temp_df.columns, aliases)
                if match:
                    score += 1

            if score > best_score:
                best_score = score
                best_df = temp_df
                best_header = header_row

        except Exception:
            continue

    if best_score > 0:
        return best_df, best_header

    return None, None


def clean_numeric(series):
    return pd.to_numeric(
        series.astype(str)
        .str.replace(",", "", regex=False)
        .str.replace("₹", "", regex=False)
        .str.replace("INR", "", regex=False)
        .str.strip(),
        errors="coerce"
    ).fillna(0)


def to_excel(df):
    output = BytesIO()
    with pd.ExcelWriter(output, engine="openpyxl") as writer:
        df.to_excel(writer, index=False, sheet_name="Reconciliation_Result")
    output.seek(0)
    return output.getvalue()


# =========================
# Column Aliases
# =========================

DEPOSIT_ALIASES = [
    "Deposit Amt (INR)",
    "Deposit Amount (INR)",
    "Deposit Amt",
    "Deposit Amount",
    "Deposit",
    "Deposits",
    "Credit Amount",
    "Cr Amount"
]

WITHDRAWAL_ALIASES = [
    "Withdrawal Amt (INR)",
    "Withdrawal Amount (INR)",
    "Withdrawal Amt",
    "Withdrawal Amount",
    "Withdrawal",
    "Withdrawals",
    "Debit Amount",
    "Dr Amount"
]

DEBIT_ALIASES = [
    "Debit (LC)",
    "Debit(LC)",
    "Debit LC",
    "Debit Amount",
    "Debit",
    "Dr",
    "Dr Amount"
]

CREDIT_ALIASES = [
    "Credit (LC)",
    "Credit(LC)",
    "Credit LC",
    "Credit Amount",
    "Credit",
    "Cr",
    "Cr Amount"
]


# =========================
# Upload Section
# =========================

st.markdown(
    '<div class="info-card"><div class="section-title">📤 Upload File</div></div>',
    unsafe_allow_html=True
)

uploaded_file = st.file_uploader("Upload Excel File", type=["xlsx"])

comparison_mode = st.selectbox(
    "Select Comparison Type",
    ["Deposit vs Debit", "Withdrawal vs Credit"]
)

if uploaded_file:
    try:
        file_bytes = uploaded_file.read()

        excel_file = pd.ExcelFile(BytesIO(file_bytes))
        sheet_names = excel_file.sheet_names

        st.info(f"Detected sheets: {sheet_names}")

        col_a, col_b = st.columns(2)

        with col_a:
            bank_sheet_name = st.selectbox(
                "Select Bank Statement Sheet",
                sheet_names,
                index=0
            )

        with col_b:
            erp_sheet_name = st.selectbox(
                "Select ERP / Book Sheet",
                sheet_names,
                index=1 if len(sheet_names) > 1 else 0
            )

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

        sheet1, sheet1_header = read_sheet_with_header_detection(
            file_bytes=file_bytes,
            sheet_name=bank_sheet_name,
            alias_groups=[left_aliases],
            max_header_rows=60
        )

        sheet2, sheet2_header = read_sheet_with_header_detection(
            file_bytes=file_bytes,
            sheet_name=erp_sheet_name,
            alias_groups=[right_aliases],
            max_header_rows=60
        )

        if sheet1 is None:
            st.error("❌ Could not detect the correct header row in Bank Statement Sheet.")
            st.stop()

        if sheet2 is None:
            st.error("❌ Could not detect the correct header row in ERP / Book Sheet.")
            st.stop()

        left_col = find_matching_column(sheet1.columns, left_aliases)
        right_col = find_matching_column(sheet2.columns, right_aliases)

        if left_col is None:
            st.error(f"❌ {left_label} column not found in selected Bank Sheet.")
            st.write("Available columns:", list(sheet1.columns))
            st.stop()

        if right_col is None:
            st.error(f"❌ {right_label} column not found in selected ERP Sheet.")
            st.write("Available columns:", list(sheet2.columns))
            st.stop()

        sheet1[left_col] = clean_numeric(sheet1[left_col])
        sheet2[right_col] = clean_numeric(sheet2[right_col])

        sheet1_filtered = sheet1[sheet1[left_col] != 0].copy()
        sheet2_filtered = sheet2[sheet2[right_col] != 0].copy()

        sheet1_filtered["Excel Row No"] = sheet1_filtered.index + sheet1_header + 2
        sheet2_filtered["Excel Row No"] = sheet2_filtered.index + sheet2_header + 2

        sheet1_filtered["_match_amount"] = sheet1_filtered[left_col].round(2)
        sheet2_filtered["_match_amount"] = sheet2_filtered[right_col].round(2)

        left_count = Counter(sheet1_filtered["_match_amount"])
        right_count = Counter(sheet2_filtered["_match_amount"])

        all_amounts = sorted(set(left_count.keys()) | set(right_count.keys()))
        result = []

        for amount in all_amounts:
            left_amt_count = left_count.get(amount, 0)
            right_amt_count = right_count.get(amount, 0)

            if left_amt_count != right_amt_count:
                left_rows = sheet1_filtered[
                    sheet1_filtered["_match_amount"] == amount
                ]["Excel Row No"].tolist()

                right_rows = sheet2_filtered[
                    sheet2_filtered["_match_amount"] == amount
                ]["Excel Row No"].tolist()

                if left_amt_count > right_amt_count:
                    missing_where = f"Missing in {right_label}"
                    missing_count = left_amt_count - right_amt_count
                else:
                    missing_where = f"Missing in {left_label}"
                    missing_count = right_amt_count - left_amt_count

                result.append({
                    "Amount": amount,
                    f"{left_label} Count": left_amt_count,
                    f"{right_label} Count": right_amt_count,
                    "Missing Where": missing_where,
                    "Missing Count": missing_count,
                    "Bank Sheet Rows": left_rows,
                    "ERP Sheet Rows": right_rows
                })

        result_df = pd.DataFrame(result)

        total_left = len(sheet1_filtered)
        total_right = len(sheet2_filtered)
        mismatch_groups = len(result_df)
        total_missing_entries = int(result_df["Missing Count"].sum()) if not result_df.empty else 0
        status_text = "Matched" if result_df.empty else "Mismatch Found"

        st.success("✅ File loaded and processed successfully")

        st.subheader("📊 Dashboard")

        c1, c2, c3, c4, c5 = st.columns(5)
        c1.metric(left_label, total_left)
        c2.metric(right_label, total_right)
        c3.metric("Mismatch Groups", mismatch_groups)
        c4.metric("Missing Entries", total_missing_entries)
        c5.metric("Status", status_text)

        tab1, tab2, tab3, tab4 = st.tabs(
            ["Summary", "Mismatches", "Preview", "Debug"]
        )

        with tab1:
            st.markdown("### Reconciliation Summary")

            summary_df = pd.DataFrame({
                "Metric": [
                    "Comparison Type",
                    "Bank Sheet",
                    "ERP Sheet",
                    "Detected Bank Header Row",
                    "Detected ERP Header Row",
                    "Bank Column",
                    "ERP Column",
                    "File Name",
                    "Mismatch Groups",
                    "Total Missing Entries"
                ],
                "Value": [
                    comparison_mode,
                    bank_sheet_name,
                    erp_sheet_name,
                    sheet1_header + 1,
                    sheet2_header + 1,
                    left_col,
                    right_col,
                    uploaded_file.name,
                    mismatch_groups,
                    total_missing_entries
                ]
            })

            st.dataframe(summary_df, use_container_width=True, hide_index=True)

            if not result_df.empty:
                chart_df = result_df[["Amount", "Missing Count"]].copy()
                chart_df["Amount"] = chart_df["Amount"].astype(str)
                st.bar_chart(chart_df.set_index("Amount"))
            else:
                st.success("✅ All amounts matched")

        with tab2:
            st.markdown("### Detailed Mismatch Report")

            if not result_df.empty:
                filter_type = st.selectbox(
                    "Filter by Missing Where",
                    ["All"] + sorted(result_df["Missing Where"].unique().tolist())
                )

                filtered_df = result_df.copy()

                if filter_type != "All":
                    filtered_df = filtered_df[
                        filtered_df["Missing Where"] == filter_type
                    ]

                st.dataframe(filtered_df, use_container_width=True)

                d1, d2 = st.columns(2)

                with d1:
                    st.download_button(
                        "⬇ Download CSV",
                        filtered_df.to_csv(index=False).encode("utf-8"),
                        "reconciliation_result.csv",
                        "text/csv"
                    )

                with d2:
                    st.download_button(
                        "⬇ Download Excel",
                        to_excel(filtered_df),
                        "reconciliation_result.xlsx",
                        "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                    )
            else:
                st.success("✅ No mismatches found")

        with tab3:
            st.markdown("### Data Preview")

            p1, p2 = st.columns(2)

            with p1:
                st.markdown("#### Bank Sheet Preview")
                st.dataframe(sheet1_filtered.head(30), use_container_width=True)

            with p2:
                st.markdown("#### ERP Sheet Preview")
                st.dataframe(sheet2_filtered.head(30), use_container_width=True)

        with tab4:
            st.markdown("### Debug Information")
            st.write("All sheets:", sheet_names)
            st.write("Selected bank sheet:", bank_sheet_name)
            st.write("Selected ERP sheet:", erp_sheet_name)
            st.write("Bank header row:", sheet1_header + 1)
            st.write("ERP header row:", sheet2_header + 1)
            st.write("Bank columns:", list(sheet1.columns))
            st.write("ERP columns:", list(sheet2.columns))
            st.write("Detected bank amount column:", left_col)
            st.write("Detected ERP amount column:", right_col)

        st.markdown("""
        <div class="footer-box">
            🚀 Built by <b>Aniruddha Pathak</b> • Smart Reconciliation Tool
        </div>
        """, unsafe_allow_html=True)

    except Exception as e:
        st.error(f"❌ Error: {e}")
