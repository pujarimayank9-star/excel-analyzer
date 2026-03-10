import streamlit as st
import pandas as pd
import matplotlib.pyplot as plt
from pptx import Presentation
from pptx.util import Inches
import tempfile

st.set_page_config(layout="wide")

st.title("Universal Excel Data Analyst")

file = st.file_uploader("Upload Excel / CSV", type=["xlsx","csv"])

if file:

    if file.name.endswith(".csv"):
        df = pd.read_csv(file)
    else:
        df = pd.read_excel(file)

    st.subheader("Dataset Preview")
    st.dataframe(df)

    numeric_cols = df.select_dtypes(include=["int64","float64"]).columns
    cat_cols = df.select_dtypes(include=["object"]).columns

    charts = []
    insights = []
    recommendations = []

    # -----------------------
    # NUMERIC ANALYSIS
    # -----------------------

    st.header("Numeric Analysis")

    for col in numeric_cols:

        avg = round(df[col].mean(),2)
        mx = df[col].max()
        mn = df[col].min()

        st.subheader(col)

        c1,c2,c3 = st.columns(3)
        c1.metric("Average",avg)
        c2.metric("Max",mx)
        c3.metric("Min",mn)

        fig,ax = plt.subplots()
        df[col].hist(ax=ax)

        st.pyplot(fig)

        charts.append((col,fig))

        insights.append(
            f"{col} ranges from {mn} to {mx} with an average of {avg}, indicating variability in the dataset."
        )

        recommendations.append(
            f"Monitor extreme values in {col} and ensure operational consistency."
        )

    # -----------------------
    # CATEGORY ANALYSIS
    # -----------------------

    st.header("Category Analysis")

    for cat in cat_cols:

        if df[cat].nunique() <= 20:

            counts = df[cat].value_counts()

            st.subheader(cat)

            fig1,ax = plt.subplots()
            counts.plot(kind="bar",ax=ax)

            st.pyplot(fig1)

            fig2,ax = plt.subplots()
            counts.plot(kind="pie",autopct="%1.1f%%",ax=ax)

            st.pyplot(fig2)

            charts.append((cat+"_bar",fig1))
            charts.append((cat+"_pie",fig2))

            best = counts.idxmax()
            worst = counts.idxmin()

            insights.append(
                f"In column {cat}, '{best}' appears most frequently while '{worst}' appears least frequently, indicating imbalance in category distribution."
            )

            recommendations.append(
                f"Analyse why '{best}' dominates in {cat} and explore strategies to improve '{worst}'."
            )

    # -----------------------
    # CORRELATION
    # -----------------------

    if len(numeric_cols) > 1:

        st.header("Correlation Analysis")

        corr = df[numeric_cols].corr()

        fig,ax = plt.subplots()
        cax = ax.matshow(corr)

        st.pyplot(fig)

        charts.append(("correlation",fig))

        insights.append(
            "Correlation analysis reveals relationships between numeric variables that may influence each other."
        )

    # -----------------------
    # INSIGHTS DISPLAY
    # -----------------------

    st.header("Insights")

    for i in insights:
        st.write("-",i)

    st.header("Recommendations")

    for r in recommendations:
        st.write("-",r)

    # -----------------------
    # PPT GENERATOR
    # -----------------------

    if st.button("Generate Professional PPT"):

        prs = Presentation()

        # Title
        slide = prs.slides.add_slide(prs.slide_layouts[0])
        slide.shapes.title.text = "Automated Data Analysis Report"
        slide.placeholders[1].text = "Generated from Excel Dataset"

        # Numeric summary
        for col in numeric_cols:

            avg = round(df[col].mean(),2)
            mx = df[col].max()
            mn = df[col].min()

            slide = prs.slides.add_slide(prs.slide_layouts[1])
            slide.shapes.title.text = f"{col} Statistics"

            slide.placeholders[1].text = f"""
Average: {avg}

Maximum: {mx}

Minimum: {mn}
"""

        # Category charts
        for cat in cat_cols:

            if df[cat].nunique() <= 20:

                counts = df[cat].value_counts()

                fig_bar,ax = plt.subplots()
                counts.plot(kind="bar",ax=ax)

                fig_pie,ax = plt.subplots()
                counts.plot(kind="pie",autopct="%1.1f%%",ax=ax)

                tmp1 = tempfile.NamedTemporaryFile(delete=False,suffix=".png")
                tmp2 = tempfile.NamedTemporaryFile(delete=False,suffix=".png")

                fig_bar.savefig(tmp1.name)
                fig_pie.savefig(tmp2.name)

                slide = prs.slides.add_slide(prs.slide_layouts[5])

                slide.shapes.title.text = cat

                slide.shapes.add_picture(tmp1.name,Inches(0.5),Inches(1.5),height=Inches(4))
                slide.shapes.add_picture(tmp2.name,Inches(6),Inches(1.5),height=Inches(4))

        # Insights slide
        slide = prs.slides.add_slide(prs.slide_layouts[1])
        slide.shapes.title.text = "Key Insights"

        slide.placeholders[1].text = "\n".join(insights)

        # Recommendation slide
        slide = prs.slides.add_slide(prs.slide_layouts[1])
        slide.shapes.title.text = "Recommendations"

        slide.placeholders[1].text = "\n".join(recommendations)

        ppt = "analysis_report.pptx"

        prs.save(ppt)

        with open(ppt,"rb") as f:

            st.download_button(
                "Download PPT",
                f,
                file_name="analysis_report.pptx"
            )
