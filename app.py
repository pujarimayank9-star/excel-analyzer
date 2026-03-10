import streamlit as st
import pandas as pd
import numpy as np
import matplotlib.pyplot as plt
from pptx import Presentation

st.set_page_config(page_title="AI Data Analyst", layout="wide")

st.title("AI Data Analyst + Strategy Engine")

file = st.file_uploader("Upload Excel File", type=["xlsx","csv"])


def generate_ppt(insights, suggestions):

    prs = Presentation()

    slide = prs.slides.add_slide(prs.slide_layouts[0])
    slide.shapes.title.text = "AI Data Analysis Report"
    slide.placeholders[1].text = "Auto generated business analysis"

    slide = prs.slides.add_slide(prs.slide_layouts[1])
    slide.shapes.title.text = "Key Insights"
    slide.placeholders[1].text = "\n".join(insights)

    slide = prs.slides.add_slide(prs.slide_layouts[1])
    slide.shapes.title.text = "Strategic Recommendations"
    slide.placeholders[1].text = "\n".join(suggestions)

    prs.save("analysis_report.pptx")


if file:

    if file.name.endswith(".csv"):
        df = pd.read_csv(file)
    else:
        df = pd.read_excel(file)

    st.subheader("Dataset Preview")
    st.dataframe(df)

    numeric_cols = df.select_dtypes(include=np.number).columns.tolist()
    cat_cols = df.select_dtypes(include="object").columns.tolist()

    insights = []
    suggestions = []

# -------------------------
# DATASET OVERVIEW
# -------------------------

    st.header("Dataset Overview")

    c1,c2,c3 = st.columns(3)

    c1.metric("Rows", df.shape[0])
    c2.metric("Columns", df.shape[1])
    c3.metric("Numeric Columns", len(numeric_cols))

# -------------------------
# NUMERIC ANALYSIS
# -------------------------

    st.header("Numeric Analysis")

    for col in numeric_cols:

        avg = df[col].mean()
        mx = df[col].max()
        mn = df[col].min()

        c1,c2,c3 = st.columns(3)

        c1.metric(f"{col} Avg", round(avg,2))
        c2.metric(f"{col} Max", mx)
        c3.metric(f"{col} Min", mn)

        fig, ax = plt.subplots()

        df[col].plot(kind="hist", ax=ax)

        ax.set_title(f"{col} Distribution")

        st.pyplot(fig)

        insights.append(f"{col} average value is {round(avg,2)}")

# -------------------------
# CATEGORY ANALYSIS
# -------------------------

    st.header("Category Performance")

    for cat in cat_cols:

        for num in numeric_cols:

            grouped = df.groupby(cat)[num].mean()

            fig, ax = plt.subplots()

            grouped.plot(kind="bar", ax=ax)

            ax.set_title(f"{cat} vs {num}")

            st.pyplot(fig)

            best = grouped.idxmax()
            worst = grouped.idxmin()

            st.write(f"Best {cat}: {best}")
            st.write(f"Worst {cat}: {worst}")

            insights.append(f"{best} drives the highest {num} within {cat}")
            insights.append(f"{worst} significantly underperforms in {cat}")

            suggestions.append(
                f"Investigate operational practices behind {best} in {cat} "
                f"and replicate those strategies across lower performing segments like {worst}."
            )

# -------------------------
# CORRELATION
# -------------------------

    if len(numeric_cols) > 1:

        st.header("Correlation Heatmap")

        corr = df[numeric_cols].corr()

        fig, ax = plt.subplots()

        cax = ax.matshow(corr)

        fig.colorbar(cax)

        ax.set_xticks(range(len(corr.columns)))
        ax.set_xticklabels(corr.columns, rotation=90)

        ax.set_yticks(range(len(corr.columns)))
        ax.set_yticklabels(corr.columns)

        st.pyplot(fig)

        insights.append("Correlation between numeric variables detected")

# -------------------------
# TREND ANALYSIS
# -------------------------

    st.header("Trend Analysis")

    date_col = None

    for col in df.columns:

        name = col.lower()

        if "date" in name or "day" in name or "time" in name:

            try:
                df[col] = pd.to_datetime(df[col])
                date_col = col
                break
            except:
                pass

    if date_col:

        st.write(f"Detected date column: {date_col}")

        for num in numeric_cols:

            trend = df.groupby(date_col)[num].sum()

            fig, ax = plt.subplots()

            trend.plot(ax=ax)

            ax.set_title(f"{num} Trend")

            st.pyplot(fig)

        insights.append("Sales trend over time detected")

        suggestions.append(
            "Analyze seasonal spikes in the trend chart and align marketing campaigns "
            "or inventory planning to high demand periods."
        )

    else:

        st.info("No date column found. Trend analysis skipped.")

# -------------------------
# AI INSIGHTS
# -------------------------

    st.header("Business Insights")

    for i in insights:
        st.write("•", i)

# -------------------------
# STRATEGIC SUGGESTIONS
# -------------------------

    st.header("Strategic Recommendations")

    for s in suggestions:
        st.write("•", s)

# -------------------------
# PPT
# -------------------------

    if st.button("Generate PowerPoint Report"):

        generate_ppt(insights, suggestions)

        with open("analysis_report.pptx","rb") as f:

            st.download_button(
                "Download PPT",
                f,
                file_name="AI_Data_Analysis_Report.pptx"
            )
