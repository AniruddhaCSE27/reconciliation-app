import streamlit as st
import pandas as pd
from collections import Counter

st.title("💰 Deposit vs Debit Reconciliation Tool")

# ==========================================
# FILE UPLOAD
# ==========================================

uploaded_file = st.file_uploader(
    "Upload Excel File (icici.xlsx)",
    type=["xlsx"]
)

if uploaded_file:

    # Read sheets
    sheet1 = pd.read_excel(uploaded_file, sheet_name="Sheet1", header=16)
    sheet2 = pd.read_excel(uploaded_file, sheet_name="Sheet2")

    st.success("File loaded successfully ✔")

    # ==========================================
    # CLEAN DATA
    # ==========================================

    sheet1['Deposit Amt (INR)'] = (
        sheet1['Deposit Amt (INR)']
        .astype(str)
        .str.replace(',', '', regex=False)
    )

    sheet1['Deposit Amt (INR)'] = pd.to_numeric(
        sheet1['Deposit Amt (INR)'],
        errors='coerce'
    ).fillna(0)

    sheet2['Debit (LC)'] = pd.to_numeric(
        sheet2['Debit (LC)'],
        errors='coerce'
    ).fillna(0)

    # keep non-zero
    sheet1 = sheet1[sheet1['Deposit Amt (INR)'] != 0].reset_index()
    sheet2 = sheet2[sheet2['Debit (LC)'] != 0].reset_index()

    # ==========================================
    # COUNT COMPARISON
    # ==========================================

    deposit_count = Counter(sheet1['Deposit Amt (INR)'])
    debit_count = Counter(sheet2['Debit (LC)'])

    all_amounts = sorted(set(deposit_count) | set(debit_count))

    result = []

    for amt in all_amounts:

        dep = deposit_count.get(amt, 0)
        deb = debit_count.get(amt, 0)

        if dep != deb:

            deposit_rows = sheet1[
                sheet1['Deposit Amt (INR)'] == amt
            ]['index'].tolist()

            debit_rows = sheet2[
                sheet2['Debit (LC)'] == amt
            ]['index'].tolist()

            if deb > dep:
                missing = "Missing in Deposit Amt (INR)"
                missing_count = deb - dep
            else:
                missing = "Missing in Debit (LC)"
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
        st.dataframe(result_df)
    else:
        st.success("All amounts matched ✔")