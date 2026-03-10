# app.py

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

    # LOAD DATA
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
    # INSIGHT FUNCTIONS
    # --------------------------

    def histogram_insight(col, avg, mn, mx):
        templates = [
            f"{col} ranges from {mn} to {mx} with an average of {round(avg,2)}, showing the spread of values in the dataset.",
            f"The distribution of {col} varies between {mn} and {mx} with mean around {round(avg,2)}.",
            f"{col} values fluctuate between {mn} and {mx}, suggesting moderate variability."
        ]
        return random.choice(templates)

    def bar_insight(col, data):
        best = data.idxmax()
        worst = data.idxmin()

        templates = [
            f"In {col}, {best} records the highest value while {worst} appears the lowest.",
            f"{best} leads the {col} comparison whereas {worst} contributes the least.",
            f"{best} dominates {col} while {worst} has the smallest presence."
        ]

        return random.choice(templates), best, worst

    def pie_insight(col, counts):
        top = counts.idxmax()
        percent = round((counts.max()/counts.sum())*100,1)

        templates = [
            f"{top} represents the largest share in {col} contributing about {percent}%.",
            f"{top} dominates the distribution of {col} with roughly {percent}% share.",
            f"{top} accounts for the highest proportion of {col}."
        ]

        return random.choice(templates)

    def recommendation(best, worst, col):
        templates = [
            f"Investigate why {best} performs strongly in {col} and replicate similar strategies.",
            f"Focus on improving the performance of {worst} within {col}.",
            f"Practices used by {best} could help improve other segments in {col}."
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

            # BAR CHART
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

            # PIE CHART
            fig2, ax = plt.subplots()
            counts.plot(kind="pie", autopct="%1.1f%%", ax=ax)

            st.pyplot(fig2)

            pie_i = pie_insight(cat, counts)

            st.write("Insight:", pie_i)

            insights.append(pie_i)

            charts.append((cat+"_pie", fig2))

    # --------------------------
    # CORRELATION ANALYSIS
    # --------------------------

    if len(numeric_cols) > 1:

        st.header("Correlation Analysis")

        corr = df[numeric_cols].corr()

        fig, ax = plt.subplots(figsize=(6,5))

        cax = ax.imshow(corr, cmap="coolwarm")

        ax.set_xticks(range(len(corr.columns)))
        ax.set_yticks(range(len(corr.columns)))

        ax.set_xticklabels(corr.columns, rotation=45)
        ax.set_yticklabels(corr.columns)

        plt.colorbar(cax)

        for i in range(len(corr.columns)):
            for j in range(len(corr.columns)):
                ax.text(j, i, round(corr.iloc[i,j],2),
                        ha="center", va="center", color="black")

        st.pyplot(fig)

        charts.append(("correlation", fig))

        st.subheader("Correlation Insights")

        for i in range(len(corr.columns)):
            for j in range(i+1, len(corr.columns)):

                col1 = corr.columns[i]
                col2 = corr.columns[j]
                value = corr.iloc[i,j]

                if value > 0.7:

                    insight = f"There is a strong positive relationship between {col1} and {col2} (correlation {round(value,2)}). When {col1} increases, {col2} also tends to increase."

                elif value > 0.4:

                    insight = f"{col1} and {col2} show a moderate relationship ({round(value,2)}), indicating some connection between these variables."

                elif value < -0.7:

                    insight = f"{col1} and {col2} have a strong negative relationship ({round(value,2)}). When {col1} increases, {col2} tends to decrease."

                else:
                    continue

                st.write("Insight:", insight)

                insights.append(insight)

    # --------------------------
    # PPT GENERATION
    # --------------------------

    if st.button("Generate PPT Report"):

        prs = Presentation()

        slide = prs.slides.add_slide(prs.slide_layouts[0])
        slide.shapes.title.text = "Automated Data Analysis Report"
        slide.placeholders[1].text = "Generated from Excel Dataset"

        i = 0

        while i < len(charts):

            name, chart1 = charts[i]

            tmp1 = tempfile.NamedTemporaryFile(delete=False, suffix=".png")
            chart1.savefig(tmp1.name)

            chart2 = None

            if i+1 < len(charts) and "pie" in charts[i+1][0]:
                chart2 = charts[i+1][1]

            slide = prs.slides.add_slide(prs.slide_layouts[5])
            slide.shapes.title.text = name.replace("_bar","").replace("_pie","")

            slide.shapes.add_picture(
                tmp1.name,
                Inches(0.5),
                Inches(1.5),
                height=Inches(3)
            )

            if chart2:

                tmp2 = tempfile.NamedTemporaryFile(delete=False, suffix=".png")
                chart2.savefig(tmp2.name)

                slide.shapes.add_picture(
                    tmp2.name,
                    Inches(5),
                    Inches(1.5),
                    height=Inches(3)
                )

                i += 1

            if i < len(insights):

                textbox = slide.shapes.add_textbox(
                    Inches(0.5),
                    Inches(4.8),
                    Inches(9),
                    Inches(1)
                )

                tf = textbox.text_frame
                tf.text = "Insight: " + insights[i]

            if i < len(recommendations):

                textbox2 = slide.shapes.add_textbox(
                    Inches(0.5),
                    Inches(5.6),
                    Inches(9),
                    Inches(1)
                )

                tf2 = textbox2.text_frame
                tf2.text = "Recommendation: " + recommendations[i]

            i += 1

        prs.save("analysis_report.pptx")

        with open("analysis_report.pptx","rb") as f:

            st.download_button(
                "Download PPT",
                f,
                file_name="analysis_report.pptx"
            )
