import streamlit as st
import pandas as pd
import numpy as np
import matplotlib.pyplot as plt
from pptx import Presentation
from pptx.util import Inches

st.set_page_config(page_title="Universal Excel AI Analyzer", layout="wide")

st.title("🤖 AI Data Analyst + Auto PPT Generator")

uploaded_file = st.file_uploader("Upload Excel File", type=["xlsx"])


def generate_ppt(df, insights, numeric_cols):

    prs = Presentation()

    # Title slide
    slide = prs.slides.add_slide(prs.slide_layouts[0])
    slide.shapes.title.text = "AI Data Analysis Report"
    slide.placeholders[1].text = "Automatically generated report"

    # Dataset overview
    slide = prs.slides.add_slide(prs.slide_layouts[1])
    slide.shapes.title.text = "Dataset Overview"

    rows, cols = df.shape

    slide.placeholders[1].text = f"""
Rows: {rows}
Columns: {cols}
Numeric Columns: {len(numeric_cols)}
"""

    chart_images = []

    # Create charts
    for col in numeric_cols[:3]:

        fig, ax = plt.subplots()

        df[col].plot(kind="hist", ax=ax)

        ax.set_title(f"{col} Distribution")

        img_name = f"{col}_chart.png"

        plt.savefig(img_name)

        chart_images.append(img_name)

        plt.close()

    # Add chart slides
    for img in chart_images:

        slide = prs.slides.add_slide(prs.slide_layouts[5])

        slide.shapes.title.text = "Data Chart"

        slide.shapes.add_picture(img, Inches(1), Inches(1.5), width=Inches(6))

    # Insights slide
    slide = prs.slides.add_slide(prs.slide_layouts[1])

    slide.shapes.title.text = "Key Insights"

    slide.placeholders[1].text = "\n".join(insights)

    # Recommendations slide
    slide = prs.slides.add_slide(prs.slide_layouts[1])

    slide.shapes.title.text = "Recommendations"

    slide.placeholders[1].text = """
Investigate high and low values
Monitor relationships between variables
Focus on improving weak segments
"""

    prs.save("report.pptx")


if uploaded_file:

    df = pd.read_excel(uploaded_file)

    st.subheader("Dataset Preview")
    st.dataframe(df)

    rows, cols = df.shape

    numeric_cols = df.select_dtypes(include=np.number).columns
    cat_cols = df.select_dtypes(include="object").columns

    st.subheader("Dataset Summary")

    c1, c2, c3 = st.columns(3)

    c1.metric("Rows", rows)
    c2.metric("Columns", cols)
    c3.metric("Numeric Columns", len(numeric_cols))

    insights = []

    # Numeric metrics
    if len(numeric_cols) > 0:

        st.subheader("Numeric Metrics")

        for col in numeric_cols:

            c1, c2, c3 = st.columns(3)

            c1.metric(f"{col} Avg", round(df[col].mean(),2))
            c2.metric(f"{col} Max", df[col].max())
            c3.metric(f"{col} Min", df[col].min())

            insights.append(f"{col} average is {round(df[col].mean(),2)}")

    # Distribution charts
    if len(numeric_cols) > 0:

        st.subheader("Distribution Charts")

        for col in numeric_cols:

            fig, ax = plt.subplots()

            df[col].plot(kind="hist", ax=ax)

            ax.set_title(f"{col} Distribution")

            st.pyplot(fig)

    # Category comparisons
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

    # Correlation heatmap
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

    # AI insights
    st.subheader("AI Insights")

    for i in insights:
        st.write("•", i)

    # Generate PPT
    if st.button("Generate PowerPoint Report"):

        generate_ppt(df, insights, numeric_cols)

        with open("report.pptx","rb") as f:

            st.download_button(
                "Download PPT Report",
                f,
                file_name="data_analysis_report.pptx"
            )
