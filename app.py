import streamlit as st
import pandas as pd
import matplotlib.pyplot as plt

st.set_page_config(page_title="AI Data Analyzer", layout="wide")

st.title("AI Data Analyzer")

file = st.file_uploader("Upload Excel or CSV", type=["xlsx","csv"])

if file:

    if file.name.endswith(".csv"):
        df = pd.read_csv(file)
    else:
        df = pd.read_excel(file)

    st.subheader("Dataset Preview")
    st.dataframe(df)

    # -------------------------
    # CLEAN DATA
    # -------------------------

    # Remove completely empty columns
    df = df.dropna(axis=1, how="all")

    # Try converting everything to numeric
    numeric_cols = []

    for col in df.columns:

        converted = pd.to_numeric(df[col], errors="coerce")

        # If at least 1 numeric value found
        if converted.notna().sum() > 0:

            df[col] = converted
            numeric_cols.append(col)

    # Remove columns which are fully NaN after conversion
    df = df.dropna(axis=1, how="all")

    if len(numeric_cols) == 0:
        st.error("No numeric column detected for analysis.")
        st.stop()

    target = numeric_cols[0]

    # -------------------------
    # BASIC STATS
    # -------------------------

    st.header("Basic Statistics")

    col1, col2, col3 = st.columns(3)

    col1.metric("Average", round(df[target].mean(),2))
    col2.metric("Max", df[target].max())
    col3.metric("Min", df[target].min())

    # -------------------------
    # DISTRIBUTION HISTOGRAM
    # -------------------------

    st.header("Distribution")

    fig, ax = plt.subplots()

    df[target].hist(ax=ax)

    st.pyplot(fig)

    # -------------------------
    # CATEGORY ANALYSIS
    # -------------------------

    st.header("Category Analysis")

    cat_cols = df.select_dtypes(include=["object"]).columns

    for col in cat_cols:

        if df[col].nunique() < 20:

            grp = df.groupby(col)[target].mean()

            fig, ax = plt.subplots()

            grp.plot(kind="bar", ax=ax)

            ax.set_title(col)

            st.pyplot(fig)

    # -------------------------
    # PIE DISTRIBUTION
    # -------------------------

    st.header("Category Share")

    for col in cat_cols:

        if df[col].nunique() < 10:

            counts = df[col].value_counts()

            fig, ax = plt.subplots()

            counts.plot(kind="pie", autopct="%1.1f%%", ax=ax)

            ax.set_ylabel("")

            st.pyplot(fig)
