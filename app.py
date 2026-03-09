import streamlit as st
import pandas as pd

st.title("Excel Sales Analyzer")

file = st.file_uploader("Upload Excel File", type=["xlsx"])

if file is not None:

    df = pd.read_excel(file)

    # Store sales
    store_sales = df.groupby("Store")["Weekly_Sales"].sum()

    st.subheader("Store Sales")
    st.bar_chart(store_sales)

    # Monthly trend
    df["Date"] = pd.to_datetime(df["Date"])
    monthly_sales = df.groupby(df["Date"].dt.to_period("M"))["Weekly_Sales"].sum()

    st.subheader("Monthly Sales Trend")
    st.line_chart(monthly_sales)

    # Insights
    st.subheader("Insights")

    top_store = store_sales.idxmax()
    best_month = monthly_sales.idxmax()

    st.write(f"Top performing store: {top_store}")
    st.write(f"Best sales month: {best_month}")
