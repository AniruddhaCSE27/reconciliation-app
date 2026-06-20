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
.block-container {
    padding-top: 1.5rem;
    padding-bottom: 2rem;
    max-width: 1450px;
}
.hero-card {
    background: linear-gradient(135deg, #0f172a, #1d4ed8);
    padding: 30px 34px;
    border-radius: 22px;
    color: white;
    box-shadow: 0 10px 25px rgba(0,0,0,0.15);
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
.info-card {
    background: white;
    padding: 18px 20px;
    border-radius: 16px;
    border: 1px solid #e2e8f0;
    margin-bottom: 0.75rem;
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

st.markdown("""
<div class="hero-card">
    <div class="hero-title">💰 Smart Bank Reconciliation Tool</div>
    <div class="hero-subtitle">
        Upload Excel files, auto-detect headers and columns, compare transactions,
        and generate a professional mismatch report with exact row references.
    </div>
</div>
""", unsafe_allow_html=True)

with st.sidebar:
    st.markdown("## ⚙️ How to Use")
    st.write("1. Upload the Excel file")
    st.write("2. Select Bank and ERP sheets")
    st.write("3. Choose comparison mode")
    st.write("4. Review mismatch rows")
    st.write("5. Download CSV or Excel report")

    st.markdown("---")
    st.markdown("## ✅ Supported Modes")
    st.write("- Deposit vs Debit")
    st.write("- Withdrawal vs Credit")

    st.markdown("---")
    st.markdown("## ✅ Smart Features")
    st.write("- Auto sheet detection")
    st.write("- Auto header detection")
    st.write("- Strict column matching")
    st.write("- Exact row-level mismatch tracing")
    st.write("- CSV and Excel export")


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
            if col_norm.startswith(alias_norm):
                return col

    return None


def read_sheet_with_header_detection(file_bytes, sheet_name, alias_groups, max_header_rows=70):
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
                if match is not None:
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
        .str.replace("Dr", "", regex=False)
        .str.replace("Cr", "", regex=False)
        .str.strip(),
        errors="coerce"
    ).fillna(0)


def to_excel(df):
    output = BytesIO()

    with pd.ExcelWriter(output, engine="openpyxl") as writer:
        df.to_excel(writer, index=False, sheet_name="Mismatch_Report")

    output.seek(0)
    return output.getvalue()


DEPOSIT_ALIASES = [
    "Deposit Amt (INR)",
    "Deposit Amount (INR)",
    "Deposit Amt",
    "Deposit Amount",
    "Deposit",
    "Deposits"
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
    "Debit",
    "Dr Amount",
    "Dr"
]

CREDIT_ALIASES = [
    "Credit (LC)",
    "Credit(LC)",
    "Credit LC",
    "Credit Amount",
    "Credit",
    "Cr Amount",
    "Cr"
]


st.markdown(
    '<div class="info-card"><b>📤 Upload File</b></div>',
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
            sheet1_name = st.selectbox(
                "Select Bank Sheet",
                sheet_names,
                index=0
            )

        with col_b:
            sheet2_name = st.selectbox(
                "Select ERP / Book Sheet",
                sheet_names,
                index=1 if len(sheet_names) > 1 else 0
            )

        if sheet1_name == sheet2_name:
            st.warning(
                "⚠️ You selected the same sheet for both sides. "
                "For correct reconciliation, select different Bank and ERP sheets."
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
            sheet_name=sheet1_name,
            alias_groups=[left_aliases],
            max_header_rows=70
        )

        sheet2, sheet2_header = read_sheet_with_header_detection(
            file_bytes=file_bytes,
            sheet_name=sheet2_name,
            alias_groups=[right_aliases],
            max_header_rows=70
        )

        if sheet1 is None:
            st.error("❌ Could not detect the correct header row in Bank Sheet.")
            st.stop()

        if sheet2 is None:
            st.error("❌ Could not detect the correct header row in ERP / Book Sheet.")
            st.stop()

        left_col = find_matching_column(sheet1.columns, left_aliases)
        right_col = find_matching_column(sheet2.columns, right_aliases)

        if left_col is None:
            st.error(f"❌ {left_label} column not found in Bank Sheet.")
            st.write("Available Bank Sheet columns:", list(sheet1.columns))
            st.stop()

        if right_col is None:
            st.error(f"❌ {right_label} column not found in ERP / Book Sheet.")
            st.write("Available ERP Sheet columns:", list(sheet2.columns))
            st.stop()

        sheet1[left_col] = clean_numeric(sheet1[left_col])
        sheet2[right_col] = clean_numeric(sheet2[right_col])

        sheet1_filtered = sheet1[sheet1[left_col] != 0].reset_index()
        sheet2_filtered = sheet2[sheet2[right_col] != 0].reset_index()

        sheet1_filtered["Excel Row No"] = sheet1_filtered["index"] + sheet1_header + 2
        sheet2_filtered["Excel Row No"] = sheet2_filtered["index"] + sheet2_header + 2

        sheet1_filtered["_match_amount"] = sheet1_filtered[left_col].round(2)
        sheet2_filtered["_match_amount"] = sheet2_filtered[right_col].round(2)

        left_count = Counter(sheet1_filtered["_match_amount"])
        right_count = Counter(sheet2_filtered["_match_amount"])

        all_amounts = sorted(set(left_count.keys()) | set(right_count.keys()))

        result = []

        for amt in all_amounts:
            left_amt_count = left_count.get(amt, 0)
            right_amt_count = right_count.get(amt, 0)

            if left_amt_count != right_amt_count:
                left_rows = sheet1_filtered[
                    sheet1_filtered["_match_amount"] == amt
                ]["Excel Row No"].tolist()

                right_rows = sheet2_filtered[
                    sheet2_filtered["_match_amount"] == amt
                ]["Excel Row No"].tolist()

                matched_count = min(left_amt_count, right_amt_count)

                extra_left_rows = left_rows[matched_count:]
                extra_right_rows = right_rows[matched_count:]

                if right_amt_count > left_amt_count:
                    missing_where = f"Missing in {left_col}"
                    missing_count = right_amt_count - left_amt_count
                    missing_from_sheet = sheet1_name
                    extra_found_in_sheet = sheet2_name
                    exact_problem_rows = extra_right_rows
                else:
                    missing_where = f"Missing in {right_col}"
                    missing_count = left_amt_count - right_amt_count
                    missing_from_sheet = sheet2_name
                    extra_found_in_sheet = sheet1_name
                    exact_problem_rows = extra_left_rows

                result.append({
                    "Amount": amt,
                    f"{left_label} Count": left_amt_count,
                    f"{right_label} Count": right_amt_count,
                    "Missing Where": missing_where,
                    "Missing Count": missing_count,
                    "Missing From Sheet": missing_from_sheet,
                    "Extra Found In Sheet": extra_found_in_sheet,
                    "Exact Problem Row Numbers": exact_problem_rows,
                    f"{sheet1_name} Rows": left_rows,
                    f"{sheet2_name} Rows": right_rows
                })

        result_df = pd.DataFrame(result)

        if not result_df.empty:
            result_df = result_df.sort_values(
                by=["Missing Count", "Amount"],
                ascending=[False, True]
            ).reset_index(drop=True)

        total_left = len(sheet1_filtered)
        total_right = len(sheet2_filtered)
        mismatch_groups = len(result_df)
        total_missing_entries = (
            int(result_df["Missing Count"].sum())
            if not result_df.empty
            else 0
        )

        status_text = "OK" if result_df.empty else "Check"

        st.success("✅ File loaded and processed successfully")
        st.subheader("📊 Dashboard")

        c1, c2, c3, c4, c5 = st.columns(5)
        c1.metric(left_label, total_left)
        c2.metric(right_label, total_right)
        c3.metric("Mismatch Groups", mismatch_groups)
        c4.metric("Total Missing Entries", total_missing_entries)
        c5.metric("Status", status_text)

        tab1, tab2, tab3, tab4 = st.tabs(
            ["Summary", "Mismatches", "Preview", "Debug"]
        )

        with tab1:
            st.markdown("### Reconciliation Overview")

            a, b = st.columns([1.35, 1])

            with a:
                if not result_df.empty:
                    chart_df = result_df[["Amount", "Missing Count"]].copy()
                    chart_df["Amount"] = chart_df["Amount"].astype(str)
                    st.bar_chart(chart_df.set_index("Amount"))
                else:
                    st.success("✅ All amounts matched")

            with b:
                summary_df = pd.DataFrame({
                    "Metric": [
                        "Comparison Type",
                        "Bank Sheet",
                        "ERP / Book Sheet",
                        "Detected Bank Header Row",
                        "Detected ERP Header Row",
                        "Detected Bank Column",
                        "Detected ERP Column",
                        "File Name",
                        "Mismatch Groups",
                        "Total Missing Entries"
                    ],
                    "Value": [
                        comparison_mode,
                        sheet1_name,
                        sheet2_name,
                        sheet1_header + 1,
                        sheet2_header + 1,
                        left_col,
                        right_col,
                        uploaded_file.name,
                        mismatch_groups,
                        total_missing_entries
                    ]
                })

                st.dataframe(
                    summary_df,
                    use_container_width=True,
                    hide_index=True
                )

        with tab2:
            st.markdown("### Detailed Mismatch Report")

            if not result_df.empty:
                col1, col2 = st.columns(2)

                with col1:
                    filter_type = st.selectbox(
                        "Filter by Missing Where",
                        options=["All"] + sorted(
                            result_df["Missing Where"].astype(str).unique().tolist()
                        )
                    )

                with col2:
                    sort_by = st.selectbox(
                        "Sort by",
                        options=[
                            "Missing Count",
                            "Amount",
                            f"{left_label} Count",
                            f"{right_label} Count"
                        ]
                    )

                filtered_result_df = result_df.copy()

                if filter_type != "All":
                    filtered_result_df = filtered_result_df[
                        filtered_result_df["Missing Where"] == filter_type
                    ]

                filtered_result_df = filtered_result_df.sort_values(
                    by=sort_by,
                    ascending=False if sort_by == "Missing Count" else True
                ).reset_index(drop=True)

                st.dataframe(
                    filtered_result_df,
                    use_container_width=True,
                    hide_index=True
                )

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
                st.success("✅ All amounts matched. No mismatches found.")

        with tab3:
            st.markdown("### Raw Data Preview")

            p1, p2 = st.columns(2)

            with p1:
                st.markdown(f"**{sheet1_name} Preview**")
                st.dataframe(
                    sheet1_filtered.head(30),
                    use_container_width=True
                )

            with p2:
                st.markdown(f"**{sheet2_name} Preview**")
                st.dataframe(
                    sheet2_filtered.head(30),
                    use_container_width=True
                )

        with tab4:
            st.markdown("### Debug Information")
            st.write("Comparison mode:", comparison_mode)
            st.write("Available sheets:", sheet_names)
            st.write("Selected Bank Sheet:", sheet1_name)
            st.write("Selected ERP / Book Sheet:", sheet2_name)
            st.write("Detected Bank header row:", sheet1_header + 1)
            st.write("Detected ERP header row:", sheet2_header + 1)
            st.write("Detected Bank column:", left_col)
            st.write("Detected ERP column:", right_col)
            st.write("Bank Sheet columns:", list(sheet1.columns))
            st.write("ERP Sheet columns:", list(sheet2.columns))

        st.markdown("""
        <div class="footer-box">
            🚀 Built by <b>Aniruddha Pathak</b> • Smart Reconciliation Tool
        </div>
        """, unsafe_allow_html=True)

    except Exception as e:
        st.error(f"❌ Error: {e}")
