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
    st.write("3. Review summary & mismatches")
    st.write("4. Download report")

    st.markdown("---")
    st.markdown("## ✅ Features")
    st.write("- Auto header detection")
    st.write("- Smart column detection")
    st.write("- Dashboard metrics")
    st.write("- CSV & Excel export")

# ==========================================
# HELPERS
# ==========================================
def clean_columns(df):
    df.columns = df.columns.astype(str).str.strip()
    df.columns = df.columns.str.replace("\n", " ", regex=False)
    df.columns = df.columns.str.replace(r"\s+", " ", regex=True)
    return df

def read_sheet2_correctly(file_bytes):
    for i in range(6):
        try:
            temp = pd.read_excel(BytesIO(file_bytes), sheet_name="Sheet2", header=i)
            temp = clean_columns(temp)
            cols = [c.lower() for c in temp.columns]

            if any("debit" in c or "withdraw" in c for c in cols):
                return temp, i
        except:
            continue
    return None, None

def find_deposit_col(cols):
    for c in ["Deposit Amt (INR)", "Deposit Amount", "Deposit"]:
        if c in cols:
            return c
    return None

def find_debit_col(cols):
    for c in ["Debit (LC)", "Debit(LC)", "Debit", "Debit Amount"]:
        if c in cols:
            return c
    return None

def clean_numeric(series):
    return pd.to_numeric(
        series.astype(str).str.replace(",", "", regex=False),
        errors="coerce"
    ).fillna(0)

def to_excel(df):
    output = BytesIO()
    with pd.ExcelWriter(output, engine="openpyxl") as writer:
        df.to_excel(writer, index=False)
    output.seek(0)
    return output.getvalue()

# ==========================================
# UPLOAD
# ==========================================
uploaded_file = st.file_uploader("📤 Upload Excel File", type=["xlsx"])

if uploaded_file:
    try:
        file_bytes = uploaded_file.read()

        sheet1 = pd.read_excel(BytesIO(file_bytes), sheet_name="Sheet1", header=16)
        sheet1 = clean_columns(sheet1)

        sheet2, header_row = read_sheet2_correctly(file_bytes)

        if sheet2 is None:
            st.error("❌ Could not detect Sheet2 header")
            st.stop()

        deposit_col = find_deposit_col(sheet1.columns)
        debit_col = find_debit_col(sheet2.columns)

        if not deposit_col or not debit_col:
            st.error("❌ Required columns not found")
            st.stop()

        sheet1[deposit_col] = clean_numeric(sheet1[deposit_col])
        sheet2[debit_col] = clean_numeric(sheet2[debit_col])

        sheet1 = sheet1[sheet1[deposit_col] != 0].reset_index()
        sheet2 = sheet2[sheet2[debit_col] != 0].reset_index()

        deposit_count = Counter(sheet1[deposit_col])
        debit_count = Counter(sheet2[debit_col])

        all_amt = sorted(set(deposit_count) | set(debit_count))

        result = []
        for amt in all_amt:
            d1 = deposit_count.get(amt, 0)
            d2 = debit_count.get(amt, 0)

            if d1 != d2:
                result.append({
                    "Amount": amt,
                    "Deposit Count": d1,
                    "Debit Count": d2,
                    "Missing": abs(d1 - d2)
                })

        result_df = pd.DataFrame(result)

        # ==========================================
        # METRICS
        # ==========================================
        st.subheader("📊 Dashboard")

        c1, c2, c3, c4 = st.columns(4)
        c1.metric("Deposits", len(sheet1))
        c2.metric("Debits", len(sheet2))
        c3.metric("Mismatches", len(result_df))
        c4.metric("Status", "OK" if result_df.empty else "Check")

        # ==========================================
        # TABS
        # ==========================================
        tab1, tab2, tab3 = st.tabs(["Summary", "Mismatches", "Preview"])

        with tab1:
            if not result_df.empty:
                st.bar_chart(result_df.set_index("Amount"))
            else:
                st.success("All matched ✔")

        with tab2:
            st.dataframe(result_df, use_container_width=True)

            st.download_button(
                "⬇ Download CSV",
                result_df.to_csv(index=False),
                "result.csv"
            )

            st.download_button(
                "⬇ Download Excel",
                to_excel(result_df),
                "result.xlsx"
            )

        with tab3:
            st.write("Sheet1")
            st.dataframe(sheet1.head(10))
            st.write("Sheet2")
            st.dataframe(sheet2.head(10))

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
