import streamlit as st
import pandas as pd

st.set_page_config(
    page_title="Sales Analytics Dashboard",
    page_icon="📊",
    layout="wide"
)

st.title("📊 Sales Analytics Dashboard")
st.markdown("Upload your Excel file and get instant insights")

file = st.file_uploader("Upload Excel File", type=["xlsx"])

if file is not None:

    df = pd.read_excel(file)

    st.subheader("Raw Data")
    st.dataframe(df)

    # Store Sales
    store_sales = df.groupby("Store")["Weekly_Sales"].sum()

    # Monthly Trend
    df["Date"] = pd.to_datetime(df["Date"], dayfirst=True)
    df["Month"] = df["Date"].dt.month

    monthly_sales = df.groupby("Month")["Weekly_Sales"].sum()

    st.subheader("Sales Charts")

    col1, col2 = st.columns(2)

    with col1:
        st.write("Store Performance")
        st.bar_chart(store_sales)

    with col2:
        st.write("Monthly Sales Trend")
        st.line_chart(monthly_sales)

    # Insights
    top_store = store_sales.idxmax()
    top_sales = store_sales.max()

    best_month = monthly_sales.idxmax()

    st.subheader("Insights")

    c1, c2, c3 = st.columns(3)

    with c1:
        st.metric("Top Store", top_store)

    with c2:
        st.metric("Top Store Sales", top_sales)

    with c3:
        st.metric("Best Sales Month", best_month)
