import streamlit as st
import pandas as pd
from io import BytesIO
from collections import Counter

# =====================================================
# PAGE CONFIG
# =====================================================
st.set_page_config(
    page_title="Smart Bank Reconciliation Tool",
    page_icon="💰",
    layout="wide"
)

# =====================================================
# CSS
# =====================================================
st.markdown("""
<style>
.block-container {
    padding-top: 1.5rem;
    max-width: 1450px;
}
.hero-card {
    background: linear-gradient(135deg, #0f172a, #1d4ed8);
    padding: 30px 34px;
    border-radius: 22px;
    color: white;
    margin-bottom: 1rem;
}
.hero-title {
    font-size: 2.1rem;
    font-weight: 800;
}
.hero-subtitle {
    font-size: 1rem;
    opacity: 0.92;
}
.footer-box {
    margin-top: 25px;
    padding: 16px 20px;
    border-radius: 16px;
    background: white;
    text-align: center;
    font-weight: 600;
}
</style>
""", unsafe_allow_html=True)

# =====================================================
# HEADER
# =====================================================
st.markdown("""
<div class="hero-card">
    <div class="hero-title">💰 Smart Bank Reconciliation Tool</div>
    <div class="hero-subtitle">
        Upload Excel files, auto-detect sheets, headers and amount columns,
        compare transactions, and generate professional row-level mismatch reports.
    </div>
</div>
""", unsafe_allow_html=True)

# =====================================================
# SIDEBAR
# =====================================================
with st.sidebar:
    st.markdown("## ⚙️ How to Use")
    st.write("1. Upload Excel file")
    st.write("2. Select Bank and ERP sheets")
    st.write("3. Choose comparison mode")
    st.write("4. Review unmatched rows")
    st.write("5. Download client report")

    st.markdown("---")
    st.markdown("## ✅ Supported")
    st.write("- Deposit vs Debit")
    st.write("- Withdrawal vs Credit")
    st.write("- Auto header detection")
    st.write("- Row-level mismatch report")
    st.write("- Amount summary report")

# =====================================================
# HELPERS
# =====================================================
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


def clean_columns(df):
    df.columns = df.columns.astype(str).str.strip()
    df.columns = df.columns.str.replace("\n", " ", regex=False)
    df.columns = df.columns.str.replace(r"\s+", " ", regex=True)
    return df


def find_matching_column(columns, aliases):
    normalized_columns = {normalize_name(col): col for col in columns}

    for alias in aliases:
        alias_norm = normalize_name(alias)
        if alias_norm in normalized_columns:
            return normalized_columns[alias_norm]

    for col in columns:
        col_norm = normalize_name(col)
        for alias in aliases:
            alias_norm = normalize_name(alias)
            if alias_norm in col_norm:
                return col

    return None


def detect_header(file_bytes, sheet_name, alias_groups, max_rows=70):
    best_df = None
    best_header = None
    best_score = 0

    for header_row in range(max_rows):
        try:
            df = pd.read_excel(
                BytesIO(file_bytes),
                sheet_name=sheet_name,
                header=header_row
            )
            df = clean_columns(df)

            score = 0
            for aliases in alias_groups:
                if find_matching_column(df.columns, aliases):
                    score += 1

            if score > best_score:
                best_score = score
                best_df = df
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
        .str.replace("Dr", "", regex=False)
        .str.replace("Cr", "", regex=False)
        .str.strip(),
        errors="coerce"
    ).fillna(0)


def prepare_transaction_df(df, amount_col, header_row, source_name):
    temp = df.copy()
    temp[amount_col] = clean_numeric(temp[amount_col])
    temp = temp[temp[amount_col] != 0].copy()

    temp["Amount"] = temp[amount_col].round(2)
    temp["Excel Row No"] = temp.index + header_row + 2
    temp["Source"] = source_name

    return temp


def build_row_level_report(bank_df, erp_df, bank_label, erp_label):
    bank_pool = {}
    erp_pool = {}

    for _, row in bank_df.iterrows():
        amt = row["Amount"]
        bank_pool.setdefault(amt, []).append(row)

    for _, row in erp_df.iterrows():
        amt = row["Amount"]
        erp_pool.setdefault(amt, []).append(row)

    all_amounts = sorted(set(bank_pool.keys()) | set(erp_pool.keys()))
    unmatched_rows = []

    for amt in all_amounts:
        bank_rows = bank_pool.get(amt, [])
        erp_rows = erp_pool.get(amt, [])

        min_match = min(len(bank_rows), len(erp_rows))

        extra_bank = bank_rows[min_match:]
        extra_erp = erp_rows[min_match:]

        for row in extra_bank:
            unmatched_rows.append({
                "Status": f"Missing in {erp_label}",
                "Amount": amt,
                "Source": bank_label,
                "Excel Row No": row["Excel Row No"],
                "Remarks": f"{bank_label} transaction exists, but matching {erp_label} transaction not found."
            })

        for row in extra_erp:
            unmatched_rows.append({
                "Status": f"Missing in {bank_label}",
                "Amount": amt,
                "Source": erp_label,
                "Excel Row No": row["Excel Row No"],
                "Remarks": f"{erp_label} transaction exists, but matching {bank_label} transaction not found."
            })

    return pd.DataFrame(unmatched_rows)


def build_amount_summary(bank_df, erp_df, bank_label, erp_label):
    bank_count = Counter(bank_df["Amount"])
    erp_count = Counter(erp_df["Amount"])

    rows = []
    all_amounts = sorted(set(bank_count.keys()) | set(erp_count.keys()))

    for amt in all_amounts:
        b_count = bank_count.get(amt, 0)
        e_count = erp_count.get(amt, 0)

        if b_count != e_count:
            rows.append({
                "Amount": amt,
                f"{bank_label} Count": b_count,
                f"{erp_label} Count": e_count,
                "Difference": abs(b_count - e_count),
                "Status": (
                    f"Extra in {bank_label}"
                    if b_count > e_count
                    else f"Extra in {erp_label}"
                )
            })

    return pd.DataFrame(rows)


def export_excel(row_report, summary_report):
    output = BytesIO()

    with pd.ExcelWriter(output, engine="openpyxl") as writer:
        row_report.to_excel(writer, index=False, sheet_name="Row_Level_Unmatched")
        summary_report.to_excel(writer, index=False, sheet_name="Amount_Summary")

    output.seek(0)
    return output.getvalue()

# =====================================================
# ALIASES
# =====================================================
DEPOSIT_ALIASES = [
    "Deposit Amt (INR)", "Deposit Amount", "Deposit Amt",
    "Deposit", "Deposits", "Credit", "Credit Amount", "Cr Amount"
]

WITHDRAWAL_ALIASES = [
    "Withdrawal Amt (INR)", "Withdrawal Amount", "Withdrawal Amt",
    "Withdrawal", "Withdrawals", "Debit", "Debit Amount", "Dr Amount"
]

DEBIT_ALIASES = [
    "Debit (LC)", "Debit LC", "Debit", "Debit Amount",
    "Dr", "Dr Amount", "Payment", "Paid Amount"
]

CREDIT_ALIASES = [
    "Credit (LC)", "Credit LC", "Credit", "Credit Amount",
    "Cr", "Cr Amount", "Receipt", "Received Amount"
]

# =====================================================
# APP
# =====================================================
uploaded_file = st.file_uploader("📤 Upload Excel File", type=["xlsx"])

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

        col1, col2 = st.columns(2)

        with col1:
            bank_sheet = st.selectbox("Select Bank Statement Sheet", sheet_names, index=0)

        with col2:
            erp_sheet = st.selectbox(
                "Select ERP / Book Sheet",
                sheet_names,
                index=1 if len(sheet_names) > 1 else 0
            )

        if bank_sheet == erp_sheet:
            st.warning(
                "⚠️ Bank sheet and ERP sheet are the same. "
                "For real reconciliation, select two different sheets or upload a file containing both bank and ERP data."
            )

        if comparison_mode == "Deposit vs Debit":
            bank_label = "Bank Deposit"
            erp_label = "ERP Debit"
            bank_aliases = DEPOSIT_ALIASES
            erp_aliases = DEBIT_ALIASES
        else:
            bank_label = "Bank Withdrawal"
            erp_label = "ERP Credit"
            bank_aliases = WITHDRAWAL_ALIASES
            erp_aliases = CREDIT_ALIASES

        bank_df_raw, bank_header = detect_header(
            file_bytes,
            bank_sheet,
            [bank_aliases],
            max_rows=70
        )

        erp_df_raw, erp_header = detect_header(
            file_bytes,
            erp_sheet,
            [erp_aliases],
            max_rows=70
        )

        if bank_df_raw is None:
            st.error("❌ Could not detect bank sheet header.")
            st.stop()

        if erp_df_raw is None:
            st.error("❌ Could not detect ERP sheet header.")
            st.stop()

        bank_amount_col = find_matching_column(bank_df_raw.columns, bank_aliases)
        erp_amount_col = find_matching_column(erp_df_raw.columns, erp_aliases)

        if bank_amount_col is None:
            st.error(f"❌ Could not find {bank_label} column.")
            st.write("Available Bank Columns:", list(bank_df_raw.columns))
            st.stop()

        if erp_amount_col is None:
            st.error(f"❌ Could not find {erp_label} column.")
            st.write("Available ERP Columns:", list(erp_df_raw.columns))
            st.stop()

        bank_txn = prepare_transaction_df(
            bank_df_raw,
            bank_amount_col,
            bank_header,
            bank_label
        )

        erp_txn = prepare_transaction_df(
            erp_df_raw,
            erp_amount_col,
            erp_header,
            erp_label
        )

        row_report = build_row_level_report(
            bank_txn,
            erp_txn,
            bank_label,
            erp_label
        )

        summary_report = build_amount_summary(
            bank_txn,
            erp_txn,
            bank_label,
            erp_label
        )

        total_bank = len(bank_txn)
        total_erp = len(erp_txn)
        total_unmatched = len(row_report)
        matched = max(total_bank, total_erp) - total_unmatched
        match_percent = round((matched / max(total_bank, total_erp)) * 100, 2) if max(total_bank, total_erp) else 0

        st.success("✅ File processed successfully")

        st.subheader("📊 Reconciliation Dashboard")

        m1, m2, m3, m4, m5 = st.columns(5)
        m1.metric(bank_label, total_bank)
        m2.metric(erp_label, total_erp)
        m3.metric("Unmatched Rows", total_unmatched)
        m4.metric("Amount Issues", len(summary_report))
        m5.metric("Match %", f"{match_percent}%")

        tab1, tab2, tab3, tab4, tab5 = st.tabs([
            "Client Report",
            "Amount Summary",
            "Bank Preview",
            "ERP Preview",
            "Debug"
        ])

        with tab1:
            st.markdown("### ✅ Client-Friendly Row-Level Unmatched Report")

            if row_report.empty:
                st.success("🎉 All transactions matched successfully.")
            else:
                status_filter = st.selectbox(
                    "Filter Status",
                    ["All"] + sorted(row_report["Status"].unique().tolist())
                )

                filtered_row_report = row_report.copy()

                if status_filter != "All":
                    filtered_row_report = filtered_row_report[
                        filtered_row_report["Status"] == status_filter
                    ]

                st.dataframe(
                    filtered_row_report,
                    use_container_width=True,
                    hide_index=True
                )

                st.download_button(
                    "⬇ Download Client Excel Report",
                    export_excel(filtered_row_report, summary_report),
                    "client_reconciliation_report.xlsx",
                    "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                )

                st.download_button(
                    "⬇ Download Client CSV Report",
                    filtered_row_report.to_csv(index=False).encode("utf-8"),
                    "client_reconciliation_report.csv",
                    "text/csv"
                )

        with tab2:
            st.markdown("### Amount-Wise Audit Summary")

            if summary_report.empty:
                st.success("✅ No amount-level mismatch found.")
            else:
                st.dataframe(
                    summary_report,
                    use_container_width=True,
                    hide_index=True
                )

        with tab3:
            st.markdown("### Bank Statement Preview")
            st.write("Detected Bank Header Row:", bank_header + 1)
            st.write("Detected Amount Column:", bank_amount_col)
            st.dataframe(bank_txn.head(100), use_container_width=True)

        with tab4:
            st.markdown("### ERP / Book Preview")
            st.write("Detected ERP Header Row:", erp_header + 1)
            st.write("Detected Amount Column:", erp_amount_col)
            st.dataframe(erp_txn.head(100), use_container_width=True)

        with tab5:
            st.markdown("### Debug Information")
            st.write("Uploaded File:", uploaded_file.name)
            st.write("Available Sheets:", sheet_names)
            st.write("Selected Bank Sheet:", bank_sheet)
            st.write("Selected ERP Sheet:", erp_sheet)
            st.write("Comparison Mode:", comparison_mode)
            st.write("Bank Columns:", list(bank_df_raw.columns))
            st.write("ERP Columns:", list(erp_df_raw.columns))

        st.markdown("""
        <div class="footer-box">
            🚀 Built by <b>Aniruddha Pathak</b> • Smart Bank Reconciliation Tool
        </div>
        """, unsafe_allow_html=True)

    except Exception as e:
        st.error(f"❌ Error: {e}")
