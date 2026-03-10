import streamlit as st
import pandas as pd
import matplotlib.pyplot as plt
from pptx import Presentation
from pptx.util import Inches

st.set_page_config(page_title="Excel Analyzer Pro", layout="wide")

st.title("📊 Excel Data Analyzer + PowerPoint Report")

file = st.file_uploader("Upload Excel file", type=["xlsx"])


def generate_ppt(df):

    prs = Presentation()

    # Title slide
    slide = prs.slides.add_slide(prs.slide_layouts[0])
    slide.shapes.title.text = "Data Analysis Report"
    slide.placeholders[1].text = "Automatically generated insights"

    # Dataset overview
    slide = prs.slides.add_slide(prs.slide_layouts[1])
    slide.shapes.title.text = "Dataset Overview"

    overview = f"""
Rows: {df.shape[0]}
Columns: {df.shape[1]}
"""

    slide.placeholders[1].text = overview

    # Numeric columns
    numeric_cols = df.select_dtypes(include="number").columns

    insights = ""

    for col in numeric_cols:

        avg = df[col].mean()
        mx = df[col].max()
        mn = df[col].min()

        insights += f"{col} average is {round(avg,2)}. Max value {mx}, Min value {mn}.\n"

        plt.figure()
        df[col].plot(kind="bar")
        plt.title(col)
        chart = f"{col}.png"
        plt.savefig(chart)

        slide = prs.slides.add_slide(prs.slide_layouts[5])
        slide.shapes.title.text = f"{col} Analysis"
        slide.shapes.add_picture(chart, Inches(1), Inches(2), width=Inches(6))

    # Insights slide
    slide = prs.slides.add_slide(prs.slide_layouts[1])
    slide.shapes.title.text = "Key Insights"
    slide.placeholders[1].text = insights

    prs.save("analysis_report.pptx")


if file:

    df = pd.read_excel(file)

    st.subheader("Dataset Preview")
    st.dataframe(df)

    numeric_cols = df.select_dtypes(include="number").columns
    cat_cols = df.select_dtypes(include="object").columns

    st.subheader("📈 Numeric Analysis")

    for col in numeric_cols:

        c1, c2, c3 = st.columns(3)

        c1.metric("Average", round(df[col].mean(),2))
        c2.metric("Max", df[col].max())
        c3.metric("Min", df[col].min())

        fig, ax = plt.subplots()
        df[col].plot(kind="bar", ax=ax)
        st.pyplot(fig)

    if len(cat_cols) > 0 and len(numeric_cols) > 0:

        st.subheader("📊 Category Comparison")

        cat = cat_cols[0]
        num = numeric_cols[0]

        comp = df.groupby(cat)[num].sum()

        fig, ax = plt.subplots()
        comp.plot(kind="bar", ax=ax)
        st.pyplot(fig)

        best = comp.idxmax()
        worst = comp.idxmin()

        st.success(f"Best {cat}: {best}")
        st.warning(f"Lowest {cat}: {worst}")

    if len(numeric_cols) > 1:

        st.subheader("🔎 Correlation Analysis")

        corr = df[numeric_cols].corr()

        st.dataframe(corr)

        fig, ax = plt.subplots()
        cax = ax.matshow(corr)
        fig.colorbar(cax)

        ax.set_xticks(range(len(corr.columns)))
        ax.set_xticklabels(corr.columns, rotation=90)

        ax.set_yticks(range(len(corr.columns)))
        ax.set_yticklabels(corr.columns)

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

        with open("analysis_report.pptx","rb") as f:

            st.download_button(
                "Download PPT Report",
                f,
                file_name="analysis_report.pptx"
            )
