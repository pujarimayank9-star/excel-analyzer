import streamlit as st
import pandas as pd
import numpy as np
import matplotlib.pyplot as plt
from pptx import Presentation

st.set_page_config(page_title="AI Business Analyst", layout="wide")

st.title("AI Business Analyst")

file = st.file_uploader("Upload Excel file", type=["xlsx","csv"])


def generate_ppt(insights,suggestions):

    prs = Presentation()

    slide = prs.slides.add_slide(prs.slide_layouts[0])
    slide.shapes.title.text = "AI Business Analysis"
    slide.placeholders[1].text = "Automated business intelligence report"

    slide = prs.slides.add_slide(prs.slide_layouts[1])
    slide.shapes.title.text = "Key Insights"
    slide.placeholders[1].text = "\n".join(insights)

    slide = prs.slides.add_slide(prs.slide_layouts[1])
    slide.shapes.title.text = "Strategic Recommendations"
    slide.placeholders[1].text = "\n".join(suggestions)

    prs.save("business_report.pptx")


if file:

    if file.name.endswith(".csv"):
        df = pd.read_csv(file)
    else:
        df = pd.read_excel(file)

    st.header("Dataset Preview")
    st.dataframe(df)

# -------------------------
# COLUMN TYPES
# -------------------------

    numeric_cols = df.select_dtypes(include=np.number).columns.tolist()
    cat_cols = df.select_dtypes(include="object").columns.tolist()

    insights = []
    suggestions = []

# -------------------------
# DATE DETECTION
# -------------------------

    date_col = None

    for col in df.columns:

        try:
            temp = pd.to_datetime(df[col],dayfirst=True)
            if temp.notna().sum() > len(df)*0.6:
                df[col] = temp
                date_col = col
                break
        except:
            pass

# -------------------------
# BASIC STATS
# -------------------------

    st.header("Basic Statistics")

    for col in numeric_cols:

        st.subheader(col)

        c1,c2,c3 = st.columns(3)

        c1.metric("Average", round(df[col].mean(),2))
        c2.metric("Max", df[col].max())
        c3.metric("Min", df[col].min())

        fig,ax = plt.subplots()

        df[col].plot(kind="hist",ax=ax)

        st.pyplot(fig)

        insights.append(f"{col} average value is {round(df[col].mean(),2)}")

# -------------------------
# CATEGORY ANALYSIS
# -------------------------

    st.header("Category Analysis")

    for cat in cat_cols:

        for num in numeric_cols:

            grp = df.groupby(cat)[num].mean()

            fig,ax = plt.subplots()

            grp.plot(kind="bar",ax=ax)

            st.pyplot(fig)

            best = grp.idxmax()
            worst = grp.idxmin()

            insights.append(f"{best} drives the highest {num} within {cat}")
            insights.append(f"{worst} shows lowest performance in {cat}")

# -------------------------
# PIE CHARTS
# -------------------------

    st.header("Distribution Share")

    for cat in cat_cols:

        counts = df[cat].value_counts()

        fig,ax = plt.subplots()

        counts.plot(kind="pie",autopct='%1.1f%%',ax=ax)

        ax.set_ylabel("")

        st.subheader(cat)
        st.pyplot(fig)

# -------------------------
# CORRELATION
# -------------------------

    if len(numeric_cols) > 1:

        st.header("Correlation")

        corr = df[numeric_cols].corr()

        fig,ax = plt.subplots()

        cax = ax.matshow(corr)

        fig.colorbar(cax)

        ax.set_xticks(range(len(corr.columns)))
        ax.set_xticklabels(corr.columns,rotation=90)

        ax.set_yticks(range(len(corr.columns)))
        ax.set_yticklabels(corr.columns)

        st.pyplot(fig)

        insights.append("Strong relationships may exist between numeric variables")

# -------------------------
# TREND ANALYSIS
# -------------------------

    st.header("Trend Analysis")

    if date_col:

        st.write("Detected date column:",date_col)

        df = df.sort_values(date_col)

        for num in numeric_cols:

            # DAILY TREND

            daily = df.groupby(date_col)[num].sum()

            fig,ax = plt.subplots()

            daily.plot(ax=ax)

            ax.set_title("Daily Trend")

            st.pyplot(fig)

            # MONTHLY TREND

            df["month"] = df[date_col].dt.to_period("M")

            monthly = df.groupby("month")[num].sum()

            fig,ax = plt.subplots()

            monthly.plot(ax=ax)

            ax.set_title("Monthly Trend")

            st.pyplot(fig)

            # MOVING AVERAGE

            ma = daily.rolling(7).mean()

            fig,ax = plt.subplots()

            daily.plot(ax=ax,label="Actual")

            ma.plot(ax=ax,label="7-day moving avg")

            ax.legend()

            st.pyplot(fig)

            insights.append("Sales moving average indicates short term trend smoothing")

            # SEASONALITY

            df["weekday"] = df[date_col].dt.day_name()

            season = df.groupby("weekday")[num].mean()

            fig,ax = plt.subplots()

            season.plot(kind="bar",ax=ax)

            ax.set_title("Weekly Seasonality")

            st.pyplot(fig)

            best_day = season.idxmax()

            insights.append(f"Highest sales typically occur on {best_day}")

    else:

        st.info("No date column detected")

# -------------------------
# SMART BUSINESS SUGGESTIONS
# -------------------------

    if "Region" in df.columns:

        region_sales = df.groupby("Region")[numeric_cols[0]].sum()

        top_region = region_sales.idxmax()

        suggestions.append(
        f"{top_region} region contributes the largest share of revenue. "
        "Expanding inventory and marketing spend in similar demographic regions "
        "may accelerate growth."
        )

    if "Product_Category" in df.columns:

        cat_sales = df.groupby("Product_Category")[numeric_cols[0]].sum()

        weak_cat = cat_sales.idxmin()

        suggestions.append(
        f"{weak_cat} category generates the lowest sales share. "
        "Consider pricing adjustments, product bundling, or targeted promotions."
        )

    if "Manager" in df.columns:

        mgr_sales = df.groupby("Manager")[numeric_cols[0]].sum()

        best_mgr = mgr_sales.idxmax()

        suggestions.append(
        f"{best_mgr} consistently drives strong performance. "
        "Analyzing their sales strategy could reveal replicable practices."
        )

# -------------------------
# BUSINESS INSIGHTS
# -------------------------

    st.header("Business Insights")

    for i in insights:

        st.write("•",i)

# -------------------------
# RECOMMENDATIONS
# -------------------------

    st.header("Strategic Recommendations")

    for s in suggestions:

        st.write("•",s)

# -------------------------
# PPT
# -------------------------

    if st.button("Generate PowerPoint Report"):

        generate_ppt(insights,suggestions)

        with open("business_report.pptx","rb") as f:

            st.download_button(
                "Download PPT",
                f,
                file_name="business_report.pptx"
            )
