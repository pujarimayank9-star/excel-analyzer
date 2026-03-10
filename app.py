import streamlit as st
import pandas as pd
import matplotlib.pyplot as plt

st.set_page_config(page_title="AI Insight Engine", layout="wide")

st.title("AI Data Insight Engine")

file = st.file_uploader("Upload Excel File", type=["xlsx","csv"])

if file:

    if file.name.endswith(".csv"):
        df = pd.read_csv(file)
    else:
        df = pd.read_excel(file)

    st.header("Dataset Preview")
    st.dataframe(df)

# ---------------------
# COLUMN DETECTION
# ---------------------

    numeric_cols = []
    cat_cols = []
    date_col = None

    for col in df.columns:

        # try date
        parsed_date = pd.to_datetime(df[col], errors="coerce", dayfirst=True)

        if parsed_date.notna().sum() > 0.7 * len(df):
            df[col] = parsed_date
            date_col = col
            continue

        # try numeric
        converted = pd.to_numeric(df[col], errors="coerce")

        if converted.notna().sum() > 0:
            df[col] = converted
            numeric_cols.append(col)
        else:
            cat_cols.append(col)

# ---------------------
# NUMERIC CHECK
# ---------------------

    if len(numeric_cols) == 0:
        st.error("No numeric column detected for analysis.")
        st.stop()

    target = numeric_cols[0]

# ---------------------
# BASIC STATS
# ---------------------

    st.header("Basic Statistics")

    c1,c2,c3 = st.columns(3)

    c1.metric("Average", round(df[target].mean(),2))
    c2.metric("Max", df[target].max())
    c3.metric("Min", df[target].min())

# ---------------------
# SEGMENT ANALYSIS
# ---------------------

    st.header("Segment Analysis")

    for cat in cat_cols:

        if df[cat].nunique() < 20:

            grp = df.groupby(cat)[target].mean()

            fig, ax = plt.subplots()
            grp.plot(kind="bar", ax=ax)

            st.pyplot(fig)

# ---------------------
# TREND ANALYSIS
# ---------------------

    if date_col:

        st.header("Trend Analysis")

        df = df.sort_values(date_col)

        daily = df.groupby(date_col)[target].sum()

        fig,ax = plt.subplots()
        daily.plot(ax=ax)

        st.pyplot(fig)

        df["month"] = df[date_col].dt.to_period("M")

        monthly = df.groupby("month")[target].sum()

        fig,ax = plt.subplots()
        monthly.plot(ax=ax)

        st.pyplot(fig)

# ---------------------
# PIE DISTRIBUTION
# ---------------------

    st.header("Distribution")

    for cat in cat_cols:

        if df[cat].nunique() < 10:

            counts = df[cat].value_counts()

            fig, ax = plt.subplots()

            counts.plot(kind="pie", autopct="%1.1f%%", ax=ax)

            st.pyplot(fig)
