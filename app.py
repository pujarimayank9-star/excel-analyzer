import streamlit as st
import pandas as pd
import matplotlib.pyplot as plt

st.set_page_config(page_title="Excel Data Analyzer", layout="wide")

st.title("Excel Data Analyzer")

file = st.file_uploader("Upload Excel or CSV file", type=["xlsx","csv"])

if file is not None:

    # -------------------
    # READ FILE
    # -------------------

    if file.name.endswith(".csv"):
        df = pd.read_csv(file)
    else:
        df = pd.read_excel(file)

    st.subheader("Dataset Preview")
    st.dataframe(df)

    # -------------------
    # DETECT COLUMNS
    # -------------------

    numeric_cols = df.select_dtypes(include=['int64','float64']).columns
    cat_cols = df.select_dtypes(include=['object']).columns

    if len(numeric_cols) == 0:
        st.error("No numeric column found.")
        st.stop()

    target = numeric_cols[0]

    # -------------------
    # BASIC STATS
    # -------------------

    st.header("Basic Statistics")

    c1,c2,c3 = st.columns(3)

    c1.metric("Average", round(df[target].mean(),2))
    c2.metric("Maximum", df[target].max())
    c3.metric("Minimum", df[target].min())

    # -------------------
    # HISTOGRAM
    # -------------------

    st.header("Distribution")

    fig, ax = plt.subplots()

    df[target].hist(ax=ax)

    st.pyplot(fig)

    # -------------------
    # CATEGORY ANALYSIS
    # -------------------

    st.header("Category Analysis")

    for col in cat_cols:

        if df[col].nunique() < 20:

            grp = df.groupby(col)[target].mean()

            fig, ax = plt.subplots()

            grp.plot(kind="bar", ax=ax)

            ax.set_title(f"{target} by {col}")

            st.pyplot(fig)

    # -------------------
    # PIE CHARTS
    # -------------------

    st.header("Category Share")

    for col in cat_cols:

        if df[col].nunique() < 10:

            counts = df[col].value_counts()

            fig, ax = plt.subplots()

            counts.plot(kind="pie", autopct="%1.1f%%", ax=ax)

            ax.set_ylabel("")

            st.pyplot(fig)

    # -------------------
    # DATE DETECTION
    # -------------------

    date_col = None

    for col in df.columns:

        try:
            converted = pd.to_datetime(df[col], dayfirst=True)
            date_col = col
            df[col] = converted
            break
        except:
            pass

    # -------------------
    # TREND ANALYSIS
    # -------------------

    if date_col:

        st.header("Trend Analysis")

        df = df.sort_values(date_col)

        trend = df.groupby(date_col)[target].sum()

        fig, ax = plt.subplots()

        trend.plot(ax=ax)

        ax.set_title("Trend over Time")

        st.pyplot(fig)

    # -------------------
    # SIMPLE INSIGHTS
    # -------------------

    st.header("Insights")

    st.write(f"Average {target} value is {round(df[target].mean(),2)}")

    for col in cat_cols:

        if df[col].nunique() < 20:

            grp = df.groupby(col)[target].mean()

            best = grp.idxmax()
            worst = grp.idxmin()

            st.write(f"Best {col}: {best}")
            st.write(f"Worst {col}: {worst}")
