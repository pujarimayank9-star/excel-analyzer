import streamlit as st
import pandas as pd
import matplotlib.pyplot as plt
import random
from pptx import Presentation
from pptx.util import Inches
import tempfile

st.set_page_config(layout="wide")

st.title("Universal Excel Data Analyst")

file = st.file_uploader("Upload Excel / CSV", type=["xlsx","csv"])

if file:

    # --------------------------
    # LOAD DATA
    # --------------------------

    if file.name.endswith(".csv"):
        df = pd.read_csv(file)
    else:
        df = pd.read_excel(file)

    st.header("Dataset Overview")
    st.write(f"Rows: {df.shape[0]} | Columns: {df.shape[1]}")
    st.dataframe(df.head())

    numeric_cols = df.select_dtypes(include=["int64","float64"]).columns
    cat_cols = df.select_dtypes(include=["object"]).columns

    insights = []
    recommendations = []
    charts = []

    # --------------------------
    # HISTOGRAM INSIGHT
    # --------------------------

    def histogram_insight(col, avg, mn, mx):
        templates = [
            f"{col} ranges from {mn} to {mx} with an average of {round(avg,2)} indicating variability.",
            f"The distribution of {col} spans between {mn} and {mx} with mean around {round(avg,2)}.",
            f"{col} values vary from {mn} to {mx} suggesting moderate spread."
        ]
        return random.choice(templates)

    # --------------------------
    # BAR INSIGHT
    # --------------------------

    def bar_insight(col, data):
        best = data.idxmax()
        worst = data.idxmin()

        templates = [
            f"{best} shows the highest value in {col} while {worst} records the lowest.",
            f"In {col}, {best} leads performance whereas {worst} appears weakest.",
            f"{best} dominates the {col} comparison while {worst} contributes least."
        ]

        return random.choice(templates), best, worst

    # --------------------------
    # PIE INSIGHT
    # --------------------------

    def pie_insight(col, counts):
        top = counts.idxmax()
        percent = round((counts.max()/counts.sum())*100,1)

        templates = [
            f"{top} represents the largest share in {col} contributing about {percent}%.",
            f"{top} dominates the {col} distribution with roughly {percent}% share.",
            f"{top} accounts for the highest portion of {col}."
        ]

        return random.choice(templates)

    # --------------------------
    # RECOMMENDATION
    # --------------------------

    def recommendation(best, worst, col):
        templates = [
            f"Investigate why {best} performs better in {col} and replicate its strategy.",
            f"Focus on improving performance of {worst} within {col}.",
            f"Strategies used by {best} could help improve other segments in {col}."
        ]
        return random.choice(templates)

    # --------------------------
    # NUMERIC ANALYSIS
    # --------------------------

    st.header("Numeric Analysis")

    for col in numeric_cols:

        st.subheader(col)

        avg = df[col].mean()
        mx = df[col].max()
        mn = df[col].min()

        c1,c2,c3 = st.columns(3)
        c1.metric("Average", round(avg,2))
        c2.metric("Max", mx)
        c3.metric("Min", mn)

        fig, ax = plt.subplots()
        df[col].hist(ax=ax)

        st.pyplot(fig)

        insight = histogram_insight(col, avg, mn, mx)
        st.write("Insight:", insight)

        insights.append(insight)
        charts.append((col,fig))

    # --------------------------
    # CATEGORY ANALYSIS
    # --------------------------

    st.header("Category Analysis")

    for cat in cat_cols:

        if df[cat].nunique() <= 20:

            st.subheader(cat)

            counts = df[cat].value_counts()

            # BAR
            fig1, ax = plt.subplots()
            counts.plot(kind="bar", ax=ax)

            st.pyplot(fig1)

            insight, best, worst = bar_insight(cat, counts)

            st.write("Insight:", insight)

            rec = recommendation(best, worst, cat)

            st.write("Recommendation:", rec)

            insights.append(insight)
            recommendations.append(rec)

            charts.append((cat+"_bar", fig1))

            # PIE
            fig2, ax = plt.subplots()
            counts.plot(kind="pie", autopct="%1.1f%%", ax=ax)

            st.pyplot(fig2)

            pie_i = pie_insight(cat, counts)

            st.write("Insight:", pie_i)

            insights.append(pie_i)

            charts.append((cat+"_pie", fig2))

    # --------------------------
    # CORRELATION
    # --------------------------

    if len(numeric_cols) > 1:

        st.header("Correlation Analysis")

        corr = df[numeric_cols].corr()

        fig, ax = plt.subplots()

        cax = ax.matshow(corr)

        st.pyplot(fig)

        charts.append(("correlation", fig))

        for i in corr.columns:
            for j in corr.columns:

                if i != j and abs(corr.loc[i,j]) > 0.7:

                    insight = f"A strong relationship exists between {i} and {j}."
                    st.write("Insight:", insight)

                    insights.append(insight)

    # --------------------------
    # FINAL RECOMMENDATIONS
    # --------------------------

    st.header("Strategic Recommendations")

    for r in recommendations:
        st.write("-", r)

    # --------------------------
    # PPT GENERATION
    # --------------------------

    if st.button("Generate PPT Report"):

        prs = Presentation()

        slide = prs.slides.add_slide(prs.slide_layouts[0])
        slide.shapes.title.text = "Automated Data Analysis Report"
        slide.placeholders[1].text = "Generated from Excel dataset"

        # Insights slide
        slide = prs.slides.add_slide(prs.slide_layouts[1])
        slide.shapes.title.text = "Key Insights"
        slide.placeholders[1].text = "\n".join(insights)

        # Recommendations
        slide = prs.slides.add_slide(prs.slide_layouts[1])
        slide.shapes.title.text = "Recommendations"
        slide.placeholders[1].text = "\n".join(recommendations)

        # Charts
        for name, chart in charts:

            tmp = tempfile.NamedTemporaryFile(delete=False, suffix=".png")

            chart.savefig(tmp.name)

            slide = prs.slides.add_slide(prs.slide_layouts[5])
            slide.shapes.title.text = name

            slide.shapes.add_picture(
                tmp.name,
                Inches(1),
                Inches(1.5),
                height=Inches(4.5)
            )

        prs.save("analysis_report.pptx")

        with open("analysis_report.pptx","rb") as f:

            st.download_button(
                "Download PPT",
                f,
                file_name="analysis_report.pptx"
            )
