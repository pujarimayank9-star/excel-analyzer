import streamlit as st
import pandas as pd

st.title("Excel Sales Analyzer")

file = st.file_uploader("Upload Excel File", type=["xlsx"])

if file is not None:

    df = pd.read_excel(file)

    st.subheader("Raw Data")
    st.write(df)

    # ---- Store Sales ----
    if "Store" in df.columns and "Weekly_Sales" in df.columns:

        store_sales = df.groupby("Store")["Weekly_Sales"].sum()

        st.subheader("Store Sales")
        st.bar_chart(store_sales)

    else:
        st.error("Excel must contain columns: Store and Weekly_Sales")

    # ---- Monthly Trend ----
    if "Date" in df.columns:

        df["Date"] = pd.to_datetime(df["Date"], errors="coerce")

        monthly_sales = df.groupby(df["Date"].dt.month)["Weekly_Sales"].sum()

        st.subheader("Monthly Sales Trend")
        st.line_chart(monthly_sales)

    else:
        st.warning("Date column not found. Monthly trend skipped.")

    # ---- Insights ----
    st.subheader("Insights")

    if "Store" in df.columns and "Weekly_Sales" in df.columns:

        top_store = store_sales.idxmax()
        top_sales = store_sales.max()

        st.write(f"Top performing store: {top_store}")
        st.write(f"Total sales of top store: {top_sales}")

    if "Date" in df.columns:

        best_month = monthly_sales.idxmax()

        st.write(f"Best sales month: {best_month}")
