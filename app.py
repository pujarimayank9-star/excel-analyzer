import streamlit as st
import pandas as pd
import numpy as np
import matplotlib.pyplot as plt
from pptx import Presentation
from pptx.util import Inches, Pt
from pptx.dml.color import RGBColor

st.set_page_config(page_title="AI Data Analyst", layout="wide")

st.title("📊 AI Data Analyst + Auto Presentation Generator")

file = st.file_uploader("Upload Excel File", type=["xlsx","csv"])


def generate_ppt(df, insights, numeric_cols):

    prs = Presentation()

    # Title slide
    slide = prs.slides.add_slide(prs.slide_layouts[0])
    slide.shapes.title.text = "AI Data Analysis Report"
    slide.placeholders[1].text = "Automatically generated analytics"

    # Dataset overview
    slide = prs.slides.add_slide(prs.slide_layouts[1])
    slide.shapes.title.text = "Dataset Overview"

    rows, cols = df.shape

    slide.placeholders[1].text = f"""
Rows: {rows}
Columns: {cols}
Numeric columns: {len(numeric_cols)}
"""

    chart_images = []

    for col in numeric_cols[:4]:

        fig, ax = plt.subplots()

        df[col].plot(kind="hist", ax=ax)

        ax.set_title(f"{col} Distribution")

        img = f"{col}.png"

        plt.savefig(img)

        chart_images.append(img)

        plt.close()

    # Chart slides
    for img in chart_images:

        slide = prs.slides.add_slide(prs.slide_layouts[5])

        title = slide.shapes.title
        title.text = "Chart Analysis"

        slide.shapes.add_picture(img, Inches(1), Inches(1.5), width=Inches(6))

    # Insights slide
    slide = prs.slides.add_slide(prs.slide_layouts[1])

    slide.shapes.title.text = "Key Insights"

    slide.placeholders[1].text = "\n".join(insights)

    # Recommendation slide
    slide = prs.slides.add_slide(prs.slide_layouts[1])

    slide.shapes.title.text = "Recommendations"

    slide.placeholders[1].text = """
Investigate patterns in high and low values.
Improve weak performing areas.
Monitor correlations between variables.
"""

    prs.save("analysis_report.pptx")


if file:

    if file.name.endswith(".csv"):
        df = pd.read_csv(file)
    else:
        df = pd.read_excel(file)

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

    # Numeric metrics
    if len(numeric_cols) > 0:

        st.subheader("Key Metrics")

        for col in numeric_cols:

            avg = df[col].mean()
            mx = df[col].max()
            mn = df[col].min()

            c1,c2,c3 = st.columns(3)

            c1.metric(f"{col} Avg", round(avg,2))
            c2.metric(f"{col} Max", mx)
            c3.metric(f"{col} Min", mn)

            insights.append(f"{col} average value is {round(avg,2)}")

    # Histogram charts
    if len(numeric_cols) > 0:

        st.subheader("Distribution Analysis")

        for col in numeric_cols:

            fig, ax = plt.subplots()

            df[col].plot(kind="hist", ax=ax)

            ax.set_title(col)

            st.pyplot(fig)

            st.write(f"Explanation: Histogram shows how {col} values are distributed in the dataset.")

    # Box plots
    if len(numeric_cols) > 0:

        st.subheader("Box Plot Analysis")

        for col in numeric_cols:

            fig, ax = plt.subplots()

            df.boxplot(column=col, ax=ax)

            ax.set_title(col)

            st.pyplot(fig)

            st.write(f"Explanation: Box plot helps identify outliers and spread in {col} values.")

    # Category vs numeric
    if len(cat_cols) > 0 and len(numeric_cols) > 0:

        st.subheader("Category Comparison")

        cat = cat_cols[0]
        num = numeric_cols[0]

        data = df.groupby(cat)[num].mean()

        fig, ax = plt.subplots()

        data.plot(kind="bar", ax=ax)

        st.pyplot(fig)

        best = data.idxmax()
        worst = data.idxmin()

        st.write(f"Best category based on {num}: **{best}**")
        st.write(f"Worst category based on {num}: **{worst}**")

        insights.append(f"{best} performs best while {worst} performs lowest in {num}")

    # Scatter plots
    if len(numeric_cols) > 1:

        st.subheader("Relationship Analysis")

        for i in range(len(numeric_cols)-1):

            x = numeric_cols[i]
            y = numeric_cols[i+1]

            fig, ax = plt.subplots()

            ax.scatter(df[x], df[y])

            ax.set_xlabel(x)
            ax.set_ylabel(y)

            st.pyplot(fig)

            st.write(f"Explanation: Scatter plot shows relationship between {x} and {y}")

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

        st.write("Explanation: Correlation heatmap shows strength of relationship between variables.")

    # Insights
    st.subheader("AI Insights")

    for i in insights:
        st.write("•", i)

    # PPT
    if st.button("Generate PowerPoint Report"):

        generate_ppt(df, insights, numeric_cols)

        with open("analysis_report.pptx","rb") as f:

            st.download_button(
                "Download PPT Report",
                f,
                file_name="AI_Data_Report.pptx"
            )
