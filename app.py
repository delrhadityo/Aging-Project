import streamlit as st
import pandas as pd
import numpy as np
from datetime import datetime
import os

# =====================
# Konfigurasi
# =====================
PAJAK_RATE = 0.11
DATA_FILE = "challenge.xlsx"

# =====================
# Load Data
# =====================
@st.cache_data
def load_data():
    if not os.path.exists(DATA_FILE):
        st.error(f"❌ File {DATA_FILE} tidak ditemukan.")
        st.stop()
    detail_df = pd.read_excel(DATA_FILE, sheet_name="Detail Pipeline")
    order_df = pd.read_excel(DATA_FILE, sheet_name="Order Philips")
    return detail_df, order_df

# =====================
# Hitung Aging
# =====================
def hitung_aging(df):
    df['Date Register'] = pd.to_datetime(df['Date Register'], errors='coerce')
    df['Bulan Closed'] = pd.to_datetime(df['Bulan Closed'], errors='coerce')
    today = pd.Timestamp.today()

    df['Aging Days'] = np.where(
        df['Bulan Closed'].notna(),
        (df['Bulan Closed'] - df['Date Register']).dt.days,
        (today - df['Date Register']).dt.days
    )

    def categorize(days):
        if pd.isna(days):
            return "Unknown"
        elif days <= 30:
            return "≤30 Hari"
        elif days <= 90:
            return "31–90 Hari"
        else:
            return ">90 Hari"

    df['Aging Category'] = df['Aging Days'].apply(categorize)
    return df

# =====================
# Hitung Profit
# =====================
def hitung_profit(df, order_df):
    # Bersihkan data order
    order_df = order_df.dropna(subset=['Nama Produk', 'Harga Perolehan'])
    order_df = order_df.drop_duplicates(subset=['Nama Produk'], keep='first')

    # Merge langsung agar tidak masalah perbedaan mapping
    df = df.merge(
        order_df[['Nama Produk', 'Harga Perolehan']],
        left_on='Type Product',
        right_on='Nama Produk',
        how='left'
    )

    df['Harga Netto'] = df['Harga Satuan'] / (1 + PAJAK_RATE)
    df['Total Jual Netto'] = df['Harga Netto'] * df['QTY']
    df['Total Perolehan'] = df['Harga Perolehan'] * df['QTY']
    df['Profit Value'] = df['Total Jual Netto'] - df['Total Perolehan']
    df['Profit %'] = (df['Profit Value'] / df['Total Jual Netto']) * 100
    df['Bulan Transaksi'] = pd.to_datetime(df['Date Register'], errors='coerce').dt.to_period('M').astype(str)
    return df


# =====================
# Conditional Formatting Style
# =====================
def color_aging(row):
    if pd.notna(row['Bulan Closed']):
        return ['background-color: lightgreen'] * len(row)
    elif row['Aging Days'] > 90:
        return ['color: red'] * len(row)
    return [''] * len(row)

# =====================
# Export Excel
# =====================
def export_excel(df):
    file_name = "dashboard_export.xlsx"
    df.to_excel(file_name, index=False)
    return file_name

# =====================
# Streamlit UI
# =====================
st.set_page_config(page_title="Dashboard Aging & Profit", layout="wide")
st.title("📊 Dashboard Aging Project & Profit")

# Load Data
detail_df, order_df = load_data()

# Hitung Aging & Profit
detail_df = hitung_aging(detail_df)
detail_df = hitung_profit(detail_df, order_df)

# Summary Aging
st.subheader("📆 Ringkasan Aging Project")
aging_summary = detail_df['Aging Category'].value_counts().reset_index()
aging_summary.columns = ['Aging Category', 'Jumlah Project']
st.dataframe(aging_summary, use_container_width=True)

# Detail Aging Table
st.write("**Detail Project**")
st.dataframe(detail_df.style.apply(color_aging, axis=1), use_container_width=True)

# Summary Profit
st.subheader("💰 Ringkasan Profit per Project")
profit_summary = detail_df.groupby('Project Name').agg({
    'Total Jual Netto': 'sum',
    'Total Perolehan': 'sum',
    'Profit Value': 'sum'
}).reset_index()
profit_summary['Profit %'] = (profit_summary['Profit Value'] / profit_summary['Total Jual Netto']) * 100
st.dataframe(profit_summary, use_container_width=True)

# Filter Per Bulan
st.subheader("📅 Filter Profit per Bulan Transaksi")
valid_months = sorted([m for m in detail_df['Bulan Transaksi'].unique() if m != 'NaT'])
if valid_months:
    bulan_terpilih = st.selectbox("Pilih Bulan", valid_months)
    df_bulan = detail_df[detail_df['Bulan Transaksi'] == bulan_terpilih]
    st.dataframe(df_bulan[['Project Name', 'Total Jual Netto', 'Total Perolehan', 'Profit Value', 'Profit %']], use_container_width=True)
else:
    st.warning("Tidak ada bulan transaksi valid.")

# Form Input
st.subheader("📝 Input Data Project Baru")
with st.form("input_form"):
    sales_name = st.text_input("Sales Name")
    date_register = st.date_input("Date Register")
    customer = st.text_input("Customer")
    project_name = st.text_input("Project Name")
    type_product = st.text_input("Type Product")
    qty = st.number_input("QTY", min_value=1, step=1)
    harga_satuan = st.number_input("Harga Satuan (Include PPN)", min_value=0.0, step=1000.0)
    submitted = st.form_submit_button("Simpan Data")
    if submitted:
        st.success("✅ Data berhasil disimpan (simulasi, belum tersimpan permanen).")

# Export Button
if st.button("💾 Export ke Excel"):
    file_path = export_excel(detail_df)
    with open(file_path, "rb") as f:
        st.download_button("Download File Excel", f, file_name=file_path)
