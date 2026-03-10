import streamlit as st
import pandas as pd
import matplotlib.pyplot as plt
from pptx import Presentation
from pptx.util import Inches

st.set_page_config(page_title="Excel Analyzer + PPT Generator", layout="wide")

st.title("📊 Excel Data Analyzer + Auto PowerPoint Report")

uploaded_file = st.file_uploader("Upload Excel File", type=["xlsx"])


def generate_ppt(df):

    prs = Presentation()

    # Title slide
    slide_layout = prs.slide_layouts[0]
    slide = prs.slides.add_slide(slide_layout)

    slide.shapes.title.text = "Excel Analysis Report"
    slide.placeholders[1].text = "Automatically generated"

    # Dataset overview
    slide_layout = prs.slide_layouts[1]
    slide = prs.slides.add_slide(slide_layout)

    slide.shapes.title.text = "Dataset Overview"

    overview = f"""
Total Rows: {df.shape[0]}
Total Columns: {df.shape[1]}
"""

    slide.placeholders[1].text = overview

    numeric_cols = df.select_dtypes(include="number").columns

    if len(numeric_cols) > 0:

        col = numeric_cols[0]

        plt.figure()
        df[col].plot(kind="hist")
        plt.title(f"{col} Distribution")

        chart_file = "chart.png"
        plt.savefig(chart_file)

        slide_layout = prs.slide_layouts[5]
        slide = prs.slides.add_slide(slide_layout)

        slide.shapes.title.text = f"{col} Distribution"

        slide.shapes.add_picture(chart_file, Inches(1), Inches(2), width=Inches(6))

    prs.save("analysis_report.pptx")


if uploaded_file:

    df = pd.read_excel(uploaded_file)

    st.subheader("Dataset Preview")
    st.dataframe(df)

    numeric_cols = df.select_dtypes(include="number").columns

    st.subheader("📈 Numeric Analysis")

    for col in numeric_cols:

        avg = df[col].mean()
        max_val = df[col].max()
        min_val = df[col].min()

        c1, c2, c3 = st.columns(3)

        with c1:
            st.metric(f"Average {col}", round(avg, 2))

        with c2:
            st.metric(f"Max {col}", max_val)

        with c3:
            st.metric(f"Min {col}", min_val)

        st.bar_chart(df[col])

    st.subheader("🔎 Correlation Analysis")

    if len(numeric_cols) > 1:

        corr = df[numeric_cols].corr()
        st.dataframe(corr)

        fig, ax = plt.subplots()
        cax = ax.matshow(corr)

        plt.xticks(range(len(corr.columns)), corr.columns, rotation=90)
        plt.yticks(range(len(corr.columns)), corr.columns)

        fig.colorbar(cax)

        st.pyplot(fig)

    st.subheader("🧠 Automatic Insights")

    for col in numeric_cols:

        avg = df[col].mean()

        if avg > df[col].max() * 0.7:
            st.write(f"{col} values are generally high.")

        elif avg < df[col].max() * 0.4:
            st.write(f"{col} values are generally low.")

        else:
            st.write(f"{col} values are moderately distributed.")

    if st.button("📥 Generate PowerPoint Report"):

        generate_ppt(df)

        with open("analysis_report.pptx", "rb") as f:
            st.download_button(
                "Download PPT Report",
                f,
                file_name="analysis_report.pptx"
            )
