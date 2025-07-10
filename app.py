
import streamlit as st
import pandas as pd
import plotly.express as px
import matplotlib.pyplot as plt
from io import BytesIO
from sklearn.linear_model import LinearRegression
from fpdf import FPDF
from datetime import datetime

# Upload dataset
st.sidebar.header("Upload Dataset")
uploaded_file = st.sidebar.file_uploader("Unggah file Excel (.xlsx) data penjualan Anda", type=["xlsx"])

@st.cache_data
def load_data(file):
    df = pd.read_excel(file, parse_dates=["Order Date", "Ship Date"])
    return df

# Load dataset
if uploaded_file:
    df = load_data(uploaded_file)
    st.success("✅ Dataset berhasil diunggah.")
    st.write("**Pratinjau Data (5 baris pertama):**")
    st.dataframe(df.head())
else:
    df = load_data("superstore_update.xlsx")
    st.info("📌 Menggunakan dataset default (superstore_update.xlsx) karena belum ada upload data.")

# Sidebar filters
st.sidebar.header("Filter Data")
segment = st.sidebar.multiselect("Segment", df["Segment"].unique(), default=df["Segment"].unique())
category = st.sidebar.multiselect("Category", df["Category"].unique(), default=df["Category"].unique())
filtered_df_temp = df[df["Segment"].isin(segment) & df["Category"].isin(category)]
product_names = sorted(filtered_df_temp["Product Name"].unique())
product_selection = st.sidebar.multiselect("Produk", product_names, default=product_names)
df_filtered = filtered_df_temp[filtered_df_temp["Product Name"].isin(product_selection)]

# KPI Section
st.title("📊 Dashboard Penjualan UMKM")
col1, col2, col3 = st.columns(3)
col1.metric("Total Penjualan", f"${df_filtered['Sales'].sum():,.2f}")
col2.metric("Total Profit", f"${df_filtered['Profit'].sum():,.2f}")
col3.metric("Jumlah Order", df_filtered["Order ID"].nunique())

# Margin Profit
st.subheader("💰 Profit Margin per Produk")
df_filtered["Profit Margin (%)"] = (df_filtered["Profit"] / df_filtered["Sales"]) * 100
margin_df = df_filtered.groupby("Product Name")[["Sales", "Profit", "Profit Margin (%)"]].mean().sort_values("Profit Margin (%)", ascending=False).head(10)
st.dataframe(margin_df.style.format({"Sales": "${:,.2f}", "Profit": "${:,.2f}", "Profit Margin (%)": "{:.2f}%"}))

# Tren Bulanan per Kategori
st.subheader("📈 Tren Penjualan Bulanan per Kategori")
monthly_category = df_filtered.groupby([pd.Grouper(key="Order Date", freq="M"), "Category"])["Sales"].sum().reset_index()
fig = px.line(monthly_category, x="Order Date", y="Sales", color="Category", markers=True,
              title="Tren Penjualan Bulanan per Kategori", labels={"Sales": "Penjualan", "Order Date": "Tanggal"})
st.plotly_chart(fig, use_container_width=True)

# Export to PNG
if st.button("📷 Export Grafik ke PNG"):
    fig.write_image("chart.png")
    with open("chart.png", "rb") as f:
        st.download_button("Unduh Grafik PNG", f, file_name="grafik_penjualan.png")

# Tren Mingguan (opsional tambahan)
st.subheader("📆 Tren Mingguan (Semua Produk)")
weekly = df_filtered.groupby([pd.Grouper(key="Order Date", freq="W-MON")])["Sales"].sum().reset_index()
fig_weekly = px.line(weekly, x="Order Date", y="Sales", title="Tren Penjualan Mingguan", markers=True)
st.plotly_chart(fig_weekly, use_container_width=True)

# Prediksi
st.subheader("📉 Prediksi Penjualan (Linear Regression)")
monthly = df_filtered.resample("M", on="Order Date")["Sales"].sum().reset_index()
monthly["MonthNum"] = range(len(monthly))
X = monthly[["MonthNum"]]
y = monthly["Sales"]
model = LinearRegression().fit(X, y)
future = pd.DataFrame({"MonthNum": [len(monthly), len(monthly)+1, len(monthly)+2]})
pred = model.predict(future)

st.write(f"📌 Prediksi Bulan Depan: **${pred[0]:,.2f}**")

# Matplotlib prediksi
fig_pred, ax = plt.subplots()
ax.plot(monthly["Order Date"], y, label="Aktual")
ax.plot(pd.date_range(start=monthly["Order Date"].iloc[-1], periods=4, freq="M")[1:], pred, linestyle="--", label="Prediksi")
ax.set_title("Prediksi Penjualan (3 Bulan ke Depan)")
ax.legend()
st.pyplot(fig_pred)

# Download Excel
def to_excel(df):
    output = BytesIO()
    with pd.ExcelWriter(output, engine="xlsxwriter") as writer:
        df.to_excel(writer, index=False, sheet_name="Laporan")
    return output.getvalue()

excel_data = to_excel(df_filtered)
st.download_button("📥 Download Laporan Excel", excel_data, file_name="laporan_usaha.xlsx")

# Download PDF
def generate_pdf(dataframe):
    pdf = FPDF()
    pdf.add_page()
    pdf.set_font("Arial", size=12)
    pdf.cell(200, 10, txt="Laporan Ringkasan", ln=True, align="C")
    pdf.cell(200, 10, txt=f"Total Penjualan: ${df_filtered['Sales'].sum():,.2f}", ln=True)
    pdf.cell(200, 10, txt=f"Total Profit: ${df_filtered['Profit'].sum():,.2f}", ln=True)
    pdf.cell(200, 10, txt=f"Tanggal: {datetime.today().strftime('%d-%m-%Y')}", ln=True)
    pdf.ln(10)
    pdf.set_font("Arial", size=10)
    for index, row in dataframe.iterrows():
        pdf.cell(200, 8, txt=f"{row['Product Name']} | Sales: ${row['Sales']:,.2f} | Profit: ${row['Profit']:,.2f}", ln=True)
    return pdf.output(dest='S').encode('latin1')

pdf_file = generate_pdf(df_filtered.head(10))
st.download_button("📄 Download PDF", pdf_file, file_name="laporan_usaha.pdf")
