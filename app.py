import streamlit as st
import pandas as pd
import matplotlib.pyplot as plt
import numpy as np
from pptx import Presentation
from pptx.util import Inches

st.set_page_config(page_title="AI Data Analyst", layout="wide")

st.title("🤖 AI Data Analyst + Auto Presentation Generator")

uploaded_file = st.file_uploader("Upload Excel File", type=["xlsx"])

def generate_ppt(df, insights):

    prs = Presentation()

    slide = prs.slides.add_slide(prs.slide_layouts[0])
    slide.shapes.title.text = "AI Data Analysis Report"
    slide.placeholders[1].text = "Automatically generated"

    slide = prs.slides.add_slide(prs.slide_layouts[1])
    slide.shapes.title.text = "Dataset Overview"

    rows, cols = df.shape

    slide.placeholders[1].text = f"""
Rows: {rows}
Columns: {cols}
"""

    slide = prs.slides.add_slide(prs.slide_layouts[1])
    slide.shapes.title.text = "AI Insights"
    slide.placeholders[1].text = "\n".join(insights)

    prs.save("report.pptx")


if uploaded_file:

    df = pd.read_excel(uploaded_file)

    st.subheader("Dataset Preview")
    st.dataframe(df)

    rows, cols = df.shape

    numeric_cols = df.select_dtypes(include=np.number).columns
    cat_cols = df.select_dtypes(include="object").columns

    st.subheader("Dataset Summary")

    c1,c2,c3 = st.columns(3)

    c1.metric("Rows", rows)
    c2.metric("Columns", cols)
    c3.metric("Numeric Columns", len(numeric_cols))

    insights = []

    if len(numeric_cols) > 0:

        st.subheader("Numeric Metrics")

        col = numeric_cols[0]

        c1,c2,c3 = st.columns(3)

        c1.metric("Average", round(df[col].mean(),2))
        c2.metric("Max", df[col].max())
        c3.metric("Min", df[col].min())

        insights.append(f"Average {col} is {round(df[col].mean(),2)}")

    # Distribution charts

    st.subheader("Distribution Analysis")

    for col in numeric_cols:

        fig, ax = plt.subplots()

        df[col].plot(kind="hist", ax=ax)

        ax.set_title(f"{col} Distribution")

        st.pyplot(fig)

    # Category vs numeric

    if len(cat_cols) > 0 and len(numeric_cols) > 0:

        st.subheader("Category Comparison")

        cat = cat_cols[0]
        num = numeric_cols[0]

        data = df.groupby(cat)[num].mean()

        fig, ax = plt.subplots()

        data.plot(kind="bar", ax=ax)

        ax.set_title(f"{num} by {cat}")

        st.pyplot(fig)

        insights.append(f"{data.idxmax()} has highest {num}")

    # Correlation

    if len(numeric_cols) > 1:

        st.subheader("Correlation Heatmap")

        corr = df[numeric_cols].corr()

        fig, ax = plt.subplots()

        cax = ax.matshow(corr)

        fig.colorbar(cax)

        ax.set_xticks(range(len(corr.columns)))
        ax.set_xticklabels(corr.columns, rotation=90)

        ax.set_yticks(range(len(corr.columns)))
        ax.set_yticklabels(corr.columns)

        st.pyplot(fig)

    # Trend analysis

    date_cols = df.select_dtypes(include=["datetime"]).columns

    if len(date_cols) > 0 and len(numeric_cols) > 0:

        st.subheader("Trend Analysis")

        date = date_cols[0]
        num = numeric_cols[0]

        df = df.sort_values(date)

        fig, ax = plt.subplots()

        ax.plot(df[date], df[num])

        st.pyplot(fig)

        insights.append(f"{num} trend analyzed over time")

    # AI insights

    st.subheader("AI Insights")

    for i in insights:

        st.write("•", i)

    # Recommendations

    st.subheader("Business Recommendations")

    if len(numeric_cols) > 0:

        st.write("• Focus on improving lower performing segments")
        st.write("• Investigate factors influencing top values")
        st.write("• Monitor trends regularly")

    # PPT

    if st.button("Generate PowerPoint Report"):

        generate_ppt(df, insights)

        with open("report.pptx","rb") as f:

            st.download_button(
                "Download PPT",
                f,
                file_name="AI_Data_Report.pptx"
            )
