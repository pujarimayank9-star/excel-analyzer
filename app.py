import streamlit as st
import pandas as pd
import numpy as np
import matplotlib.pyplot as plt

st.set_page_config(page_title="AI Insight Engine", layout="wide")

st.title("AI Business Insight Engine")

file = st.file_uploader("Upload Excel File", type=["xlsx","csv"])

if file:

    if file.name.endswith(".csv"):
        df = pd.read_csv(file)
    else:
        df = pd.read_excel(file)

    st.header("Dataset Preview")
    st.dataframe(df)

# -----------------------
# COLUMN DETECTION
# -----------------------

    date_col = None
    numeric_cols = []
    cat_cols = []

    for col in df.columns:

        parsed = pd.to_datetime(df[col], dayfirst=True, errors="coerce")

        if parsed.notna().sum() > len(df)*0.6:
            df[col] = parsed
            date_col = col
            continue

        converted = pd.to_numeric(df[col], errors="coerce")

        if converted.notna().sum() > len(df)*0.6:
            df[col] = converted
            numeric_cols.append(col)

        else:
            cat_cols.append(col)

# -----------------------
# CHECK TARGET
# -----------------------

    if not numeric_cols:

        st.error("No numeric column detected for analysis.")
        st.stop()

    target = numeric_cols[0]

# -----------------------
# BASIC STATS
# -----------------------

    st.header("Basic Statistics")

    avg = df[target].mean()
    mx = df[target].max()
    mn = df[target].min()

    c1,c2,c3 = st.columns(3)

    c1.metric("Average", round(avg,2))
    c2.metric("Max", mx)
    c3.metric("Min", mn)

# -----------------------
# DRIVER ANALYSIS
# -----------------------

    st.header("Revenue Driver Analysis")

    driver_scores = {}

    for cat in cat_cols:

        if df[cat].nunique() < 20:

            variation = df.groupby(cat)[target].mean().std()

            driver_scores[cat] = variation

    sorted_drivers = sorted(driver_scores.items(), key=lambda x:x[1], reverse=True)

    for d in sorted_drivers:
        st.write("Driver:", d[0])

# -----------------------
# SEGMENT ANALYSIS
# -----------------------

    st.header("Segment Analysis")

    insights = []
    suggestions = []

    for cat in cat_cols:

        if df[cat].nunique() < 20:

            grp = df.groupby(cat)[target].mean()

            fig, ax = plt.subplots()

            grp.plot(kind="bar", ax=ax)

            st.pyplot(fig)

            best = grp.idxmax()
            worst = grp.idxmin()

            insights.append(f"{best} performs best in {cat}")
            insights.append(f"{worst} performs weakest in {cat}")

            suggestions.append(
                f"Improving {worst} segment performance in {cat} could increase revenue."
            )

# -----------------------
# CROSS ANALYSIS
# -----------------------

    if len(cat_cols) >= 2:

        st.header("Cross Segment Analysis")

        cat1 = cat_cols[0]
        cat2 = cat_cols[1]

        pivot = df.pivot_table(values=target,index=cat1,columns=cat2,aggfunc="mean")

        st.dataframe(pivot)

# -----------------------
# TREND ANALYSIS
# -----------------------

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

        moving = daily.rolling(7).mean()

        fig,ax = plt.subplots()

        daily.plot(ax=ax,label="Daily")

        moving.plot(ax=ax,label="Moving Avg")

        ax.legend()

        st.pyplot(fig)

# -----------------------
# INSIGHTS
# -----------------------

    st.header("Business Insights")

    for i in insights:
        st.write("•",i)

# -----------------------
# SUGGESTIONS
# -----------------------

    st.header("Strategic Suggestions")

    for s in suggestions:
        st.write("•",s)
