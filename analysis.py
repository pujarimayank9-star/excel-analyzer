import streamlit as st
import pandas as pd

def dataset_overview(df):
    st.header("📊 Dataset Overview")

    c1, c2, c3 = st.columns(3)

    c1.metric("Rows", df.shape[0])
    c2.metric("Columns", df.shape[1])
    c3.metric("Missing Values", int(df.isnull().sum().sum()))

    st.dataframe(df.head())


def dataset_detection(df):
    st.header("🔍 Dataset Type")

    cols = [c.lower() for c in df.columns]

    dtype = "Generic"

    if any(x in cols for x in ["sales", "revenue", "profit"]):
        dtype = "Sales"

    elif any(x in cols for x in ["marks", "grade", "score"]):
        dtype = "Education"

    elif any(x in cols for x in ["employee", "salary"]):
        dtype = "HR"

    elif any(x in cols for x in ["product", "stock", "inventory"]):
        dtype = "Inventory"

    st.success(f"Detected Dataset: {dtype}")


def missing_analysis(df):
    st.header("❗ Missing Values")

    missing = df.isnull().sum()

    st.dataframe(missing[missing > 0])


def duplicate_analysis(df):
    st.header("📄 Duplicate Records")

    duplicates = df.duplicated().sum()

    st.metric("Duplicate Rows", duplicates)


def numeric_analysis(df):

    st.header("📈 Numeric Analysis")

    numeric = df.select_dtypes(include="number")

    if numeric.empty:
        st.info("No numeric columns found.")
        return

    st.dataframe(numeric.describe())


def category_analysis(df):

    st.header("📊 Category Analysis")

    categorical = df.select_dtypes(include="object")

    for col in categorical.columns:

        st.subheader(col)

        st.write(categorical[col].value_counts())


def correlation_analysis(df):

    st.header("🔥 Correlation")

    numeric = df.select_dtypes(include="number")

    if len(numeric.columns) < 2:
        return

    st.dataframe(numeric.corr())


def outlier_analysis(df):

    st.header("🚨 Outlier Detection")

    numeric = df.select_dtypes(include="number")

    for col in numeric.columns:

        q1 = numeric[col].quantile(0.25)
        q3 = numeric[col].quantile(0.75)

        iqr = q3 - q1

        outliers = numeric[
            (numeric[col] < q1 - 1.5 * iqr) |
            (numeric[col] > q3 + 1.5 * iqr)
        ]

        st.write(f"{col}: {len(outliers)} outliers")


def trend_analysis(df):

    st.header("📉 Trend Analysis")

    st.info("Coming in next version.")
