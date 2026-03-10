import streamlit as st
import pandas as pd
import matplotlib.pyplot as plt
from pptx import Presentation
from pptx.util import Inches
import tempfile

st.set_page_config(layout="wide")

st.title("Universal AI Data Analyst")

file = st.file_uploader("Upload Excel or CSV", type=["xlsx","csv"])

if file:

    # -------------------
    # LOAD DATA
    # -------------------

    if file.name.endswith(".csv"):
        df = pd.read_csv(file)
    else:
        df = pd.read_excel(file)

    st.subheader("Dataset Preview")
    st.dataframe(df)

    # -------------------
    # COLUMN TYPES
    # -------------------

    numeric_cols = df.select_dtypes(include=["int64","float64"]).columns
    cat_cols = df.select_dtypes(include=["object"]).columns

    st.write("Numeric columns:", list(numeric_cols))
    st.write("Category columns:", list(cat_cols))

    charts = []
    insights = []
    recommendations = []

    # -------------------
    # NUMERIC ANALYSIS
    # -------------------

    st.header("Numeric Analysis")

    for col in numeric_cols:

        st.subheader(col)

        avg = round(df[col].mean(),2)
        mx = df[col].max()
        mn = df[col].min()

        c1,c2,c3 = st.columns(3)

        c1.metric("Average",avg)
        c2.metric("Max",mx)
        c3.metric("Min",mn)

        # HISTOGRAM
        fig,ax = plt.subplots()
        df[col].hist(ax=ax)

        st.pyplot(fig)

        charts.append((col,fig))

        insight = f"""
Column {col} shows an average value of {avg}.
Maximum recorded value is {mx} while minimum is {mn}.
This indicates the spread of values across the dataset.
"""

        insights.append(insight)

        recommendations.append(
            f"Monitor extreme values in {col} and maintain consistency."
        )

    # -------------------
    # CATEGORY ANALYSIS
    # -------------------

    st.header("Category Analysis")

    for cat in cat_cols:

        if df[cat].nunique() <= 20:

            st.subheader(cat)

            counts = df[cat].value_counts()

            # BAR
            fig1,ax = plt.subplots()
            counts.plot(kind="bar",ax=ax)

            st.pyplot(fig1)

            # PIE
            fig2,ax = plt.subplots()
            counts.plot(kind="pie",autopct="%1.1f%%",ax=ax)

            st.pyplot(fig2)

            charts.append((cat,fig1))
            charts.append((cat,fig2))

            best = counts.idxmax()
            worst = counts.idxmin()

            insight = f"""
Category {best} appears most frequently in {cat}.
Category {worst} appears least frequently.
This suggests unequal distribution among categories.
"""

            insights.append(insight)

            recommendations.append(
                f"Investigate why {best} dominates and whether balance is needed."
            )

    # -------------------
    # CORRELATION
    # -------------------

    if len(numeric_cols) > 1:

        st.header("Correlation Analysis")

        corr = df[numeric_cols].corr()

        fig,ax = plt.subplots()

        cax = ax.matshow(corr)

        st.pyplot(fig)

        charts.append(("correlation",fig))

        insights.append(
            "Correlation analysis shows relationships between numeric variables."
        )

        recommendations.append(
            "Investigate strongly correlated variables for deeper insights."
        )

    # -------------------
    # INSIGHTS
    # -------------------

    st.header("Insights")

    for i in insights:
        st.write("-",i)

    # -------------------
    # RECOMMENDATIONS
    # -------------------

    st.header("Recommendations")

    for r in recommendations:
        st.write("-",r)

    # -------------------
    # PPT GENERATION
    # -------------------

    if st.button("Generate Professional PPT"):

        prs = Presentation()

        # TITLE
        slide = prs.slides.add_slide(prs.slide_layouts[0])
        slide.shapes.title.text = "Automated Data Analysis"
        slide.placeholders[1].text = "AI Generated Business Intelligence"

        # INSIGHTS
        slide = prs.slides.add_slide(prs.slide_layouts[1])
        slide.shapes.title.text = "Key Insights"

        text = "\n".join(insights)

        slide.placeholders[1].text = text

        # RECOMMENDATION
        slide = prs.slides.add_slide(prs.slide_layouts[1])
        slide.shapes.title.text = "Recommendations"

        slide.placeholders[1].text = "\n".join(recommendations)

        # CHARTS
        for name,chart in charts:

            tmp = tempfile.NamedTemporaryFile(delete=False,suffix=".png")

            chart.savefig(tmp.name)

            slide = prs.slides.add_slide(prs.slide_layouts[5])

            slide.shapes.title.text = name

            slide.shapes.add_picture(
                tmp.name,
                Inches(1),
                Inches(1.5),
                height=Inches(4.5)
            )

        ppt = "analysis_report.pptx"

        prs.save(ppt)

        with open(ppt,"rb") as f:

            st.download_button(
                "Download PPT",
                f,
                file_name="analysis_report.pptx"
            )
