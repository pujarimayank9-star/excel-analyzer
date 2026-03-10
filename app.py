import streamlit as st
import pandas as pd
import numpy as np
import matplotlib.pyplot as plt
from pptx import Presentation
from pptx.util import Inches

st.set_page_config(page_title="Universal Excel AI Analyzer", layout="wide")

st.title("🤖 Universal Excel AI Data Analyst")

file = st.file_uploader("Upload Excel File", type=["xlsx"])


def generate_ppt(df, insights):

    prs = Presentation()

    slide = prs.slides.add_slide(prs.slide_layouts[0])
    slide.shapes.title.text = "Excel Data Analysis Report"
    slide.placeholders[1].text = "Auto generated AI analysis"

    slide = prs.slides.add_slide(prs.slide_layouts[1])
    slide.shapes.title.text = "Dataset Overview"

    rows, cols = df.shape

    slide.placeholders[1].text = f"""
Rows: {rows}
Columns: {cols}
"""

    slide = prs.slides.add_slide(prs.slide_layouts[1])
    slide.shapes.title.text = "Insights"

    slide.placeholders[1].text = "\n".join(insights)

    prs.save("report.pptx")


if file:

    df = pd.read_excel(file)

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

    # numeric metrics

    if len(numeric_cols) > 0:

        st.subheader("Numeric Metrics")

        for col in numeric_cols:

            c1, c2, c3 = st.columns(3)

            c1.metric(f"{col} Avg", round(df[col].mean(),2))
            c2.metric(f"{col} Max", df[col].max())
            c3.metric(f"{col} Min", df[col].min())

            insights.append(f"{col} average is {round(df[col].mean(),2)}")

    # histograms

    if len(numeric_cols) > 0:

        st.subheader("Distribution Charts")

        for col in numeric_cols:

            fig, ax = plt.subplots()

            df[col].plot(kind="hist", ax=ax)

            ax.set_title(f"{col} Distribution")

            st.pyplot(fig)

    # box plots

    if len(numeric_cols) > 0:

        st.subheader("Box Plots")

        for col in numeric_cols:

            fig, ax = plt.subplots()

            df.boxplot(column=col, ax=ax)

            ax.set_title(f"{col} Box Plot")

            st.pyplot(fig)

    # category vs numeric

    if len(cat_cols) > 0 and len(numeric_cols) > 0:

        st.subheader("Category Comparisons")

        for cat in cat_cols:

            for num in numeric_cols:

                try:

                    data = df.groupby(cat)[num].mean()

                    fig, ax = plt.subplots()

                    data.plot(kind="bar", ax=ax)

                    ax.set_title(f"{num} by {cat}")

                    st.pyplot(fig)

                except:
                    pass

    # scatter plots

    if len(numeric_cols) > 1:

        st.subheader("Scatter Relationships")

        for i in range(len(numeric_cols)-1):

            x = numeric_cols[i]
            y = numeric_cols[i+1]

            fig, ax = plt.subplots()

            ax.scatter(df[x], df[y])

            ax.set_xlabel(x)
            ax.set_ylabel(y)

            st.pyplot(fig)

    # correlation

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

    # recommendations

    st.subheader("Recommendations")

    st.write("• Investigate high and low values in dataset")
    st.write("• Monitor correlations between variables")
    st.write("• Focus on improving weak segments")

    # PPT

    if st.button("Generate PowerPoint Report"):

        generate_ppt(df, insights)

        with open("report.pptx","rb") as f:

            st.download_button(
                "Download PPT Report",
                f,
                file_name="data_analysis_report.pptx"
            )
