import streamlit as st
import pandas as pd
from collections import Counter
from io import BytesIO

st.set_page_config(
    page_title="Deposit vs Debit Reconciliation Tool",
    page_icon="💰",
    layout="wide"
)

st.title("💰 Deposit vs Debit Reconciliation Tool")
st.caption("Upload the ICICI Excel file and compare Deposit Amount vs Debit Amount automatically.")

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
    possible_debit_keywords = ["debit", "withdraw"]

    for header_row in range(0, 6):
        try:
            temp_df = pd.read_excel(BytesIO(file_bytes), sheet_name="Sheet2", header=header_row)
            temp_df = clean_columns(temp_df)

            cols = [str(col).lower() for col in temp_df.columns]

            if any(
                keyword in col
                for col in cols
                for keyword in possible_debit_keywords
            ):
                return temp_df, header_row
        except Exception:
            continue

    return None, None

def find_debit_column(columns):
    possible_debit_cols = [
        "Debit (LC)",
        "Debit(LC)",
        "Debit",
        "Debit LC",
        "Debit Amount",
        "Withdrawal Amount",
        "Withdrawals"
    ]
    for col in possible_debit_cols:
        if col in columns:
            return col
    return None

# ==========================================
# FILE UPLOAD
# ==========================================

uploaded_file = st.file_uploader(
    "Upload Excel File (icici.xlsx)",
    type=["xlsx"]
)

if uploaded_file:
    try:
        file_bytes = uploaded_file.read()

        # ==========================================
        # READ SHEETS
        # ==========================================
        sheet1 = pd.read_excel(BytesIO(file_bytes), sheet_name="Sheet1", header=16)
        sheet1 = clean_columns(sheet1)

        sheet2, detected_header = read_sheet2_correctly(file_bytes)

        if sheet2 is None:
            st.error("Could not detect the correct header row in Sheet2.")
            st.stop()

        st.success("File loaded successfully ✔")

        # ==========================================
        # DETECT REQUIRED COLUMNS
        # ==========================================
        deposit_col = "Deposit Amt (INR)"
        debit_col = find_debit_column(sheet2.columns)

        if deposit_col not in sheet1.columns:
            st.error(f"Column '{deposit_col}' not found in Sheet1.")
            st.write("Available Sheet1 columns:", list(sheet1.columns))
            st.stop()

        if debit_col is None:
            st.error("Debit column not found in Sheet2.")
            st.write("Available Sheet2 columns:", list(sheet2.columns))
            st.stop()

        # ==========================================
        # CLEAN DATA
        # ==========================================
        sheet1[deposit_col] = (
            sheet1[deposit_col]
            .astype(str)
            .str.replace(",", "", regex=False)
            .str.strip()
        )

        sheet2[debit_col] = (
            sheet2[debit_col]
            .astype(str)
            .str.replace(",", "", regex=False)
            .str.strip()
        )

        sheet1[deposit_col] = pd.to_numeric(sheet1[deposit_col], errors="coerce").fillna(0)
        sheet2[debit_col] = pd.to_numeric(sheet2[debit_col], errors="coerce").fillna(0)

        # Keep only non-zero values
        sheet1 = sheet1[sheet1[deposit_col] != 0].reset_index()
        sheet2 = sheet2[sheet2[debit_col] != 0].reset_index()

        # ==========================================
        # COUNT COMPARISON
        # ==========================================
        deposit_count = Counter(sheet1[deposit_col])
        debit_count = Counter(sheet2[debit_col])

        all_amounts = sorted(set(deposit_count.keys()) | set(debit_count.keys()))

        result = []

        for amt in all_amounts:
            dep = deposit_count.get(amt, 0)
            deb = debit_count.get(amt, 0)

            if dep != deb:
                deposit_rows = sheet1[sheet1[deposit_col] == amt]["index"].tolist()
                debit_rows = sheet2[sheet2[debit_col] == amt]["index"].tolist()

                if deb > dep:
                    missing = f"Missing in {deposit_col}"
                    missing_count = deb - dep
                else:
                    missing = f"Missing in {debit_col}"
                    missing_count = dep - deb

                result.append({
                    "Amount": amt,
                    "Deposit Count": dep,
                    "Debit Count": deb,
                    "Missing Where": missing,
                    "Missing Count": missing_count,
                    "Sheet1 Rows": deposit_rows,
                    "Sheet2 Rows": debit_rows
                })

        # ==========================================
        # SHOW RESULT
        # ==========================================
        st.subheader("📊 Comparison Result")

        if result:
            result_df = pd.DataFrame(result)
            st.dataframe(result_df, use_container_width=True)

            csv = result_df.to_csv(index=False).encode("utf-8")
            st.download_button(
                label="⬇ Download Result as CSV",
                data=csv,
                file_name="reconciliation_result.csv",
                mime="text/csv"
            )
        else:
            st.success("All amounts matched ✔")

        # ==========================================
        # DEBUG INFO
        # ==========================================
        with st.expander("View detected columns / debug info"):
            st.write("Detected Sheet2 header row:", detected_header)
            st.write("Detected deposit column:", deposit_col)
            st.write("Detected debit column:", debit_col)
            st.write("Sheet1 columns:", list(sheet1.columns))
            st.write("Sheet2 columns:", list(sheet2.columns))

    except Exception as e:
        st.error(f"Error while processing file: {e}")
