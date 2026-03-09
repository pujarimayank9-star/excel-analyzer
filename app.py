import streamlit as st
import pandas as pd

st.title("Excel Sales Analyzer")

file = st.file_uploader("Upload Excel File")

if file:

    df = pd.read_excel(file)

    store_sales = df.groupby("Store")["Weekly_Sales"].sum()

    st.subheader("Store Sales")
    st.bar_chart(store_sales)

    df["Date"] = pd.to_datetime(df["Date"], dayfirst=True)

    monthly_sales = df.groupby(df["Date"].dt.month)["Weekly_Sales"].sum()

    st.subheader("Monthly Sales Trend")
    st.line_chart(monthly_sales)

    top_store = store_sales.idxmax()

    st.write(f"Top performing store is Store {top_store}")
