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

    # ---------------- LOAD DATA ----------------
    if file.name.endswith(".csv"):
        df = pd.read_csv(file)
    else:
        df = pd.read_excel(file)

    st.header("Dataset Overview")
    st.write(f"Rows: {df.shape[0]} | Columns: {df.shape[1]}")
    st.dataframe(df.head())

    numeric_cols = df.select_dtypes(include=["int64","float64"]).columns
    cat_cols = df.select_dtypes(include=["object"]).columns

    reports = []
    overall_insights = []
    final_recommendations = []

    # ---------------- NUMERIC ANALYSIS ----------------

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

        insight = (
            f"{col} ranges from {mn} to {mx} with an average of {round(avg,2)} "
            f"showing how values are distributed in the dataset."
        )

        st.write("Insight:", insight)

        spread = mx - mn
        if spread > avg * 0.5:
            overall_insights.append(
                f"{col} shows high variability meaning values fluctuate widely."
            )
        else:
            overall_insights.append(
                f"{col} values remain relatively stable across the dataset."
            )

        reports.append({
            "title": col,
            "chart1": fig,
            "chart2": None,
            "insight": insight
        })

    # ---------------- CATEGORY ANALYSIS ----------------

    st.header("Category Analysis")

    for cat in cat_cols:

        if df[cat].nunique() <= 20:

            st.subheader(cat)

            counts = df[cat].value_counts()

            fig1, ax = plt.subplots()
            counts.plot(kind="bar", ax=ax)
            st.pyplot(fig1)

            best = counts.idxmax()
            percent = round((counts.max()/counts.sum())*100,1)

            insight = (
                f"{best} represents the largest portion in {cat} "
                f"contributing about {percent}% of the total distribution."
            )

            st.write("Insight:", insight)

            fig2, ax = plt.subplots()
            counts.plot(kind="pie", autopct="%1.1f%%", ax=ax)
            st.pyplot(fig2)

            reports.append({
                "title": cat,
                "chart1": fig1,
                "chart2": fig2,
                "insight": insight
            })

    # ---------------- CORRELATION ANALYSIS ----------------

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

        correlation_text = ""

        for i in range(len(corr.columns)):
            for j in range(i+1, len(corr.columns)):

                col1 = corr.columns[i]
                col2 = corr.columns[j]
                value = corr.iloc[i,j]

                if abs(value) > 0.7:
                    strength = "strong"
                elif abs(value) > 0.4:
                    strength = "moderate"
                else:
                    strength = "weak"

                if value > 0:
                    relation = "positive"
                else:
                    relation = "negative"

                explanation = (
                    f"{col1} and {col2} have a {strength} {relation} relationship "
                    f"(correlation {round(value,2)})."
                )

                st.write("Insight:", explanation)
                correlation_text += explanation + " "

        reports.append({
            "title": "Correlation Analysis",
            "chart1": fig,
            "chart2": None,
            "insight": correlation_text
        })

    # ---------------- OVERALL INSIGHTS ----------------

    st.header("Overall Observations")

    for text in overall_insights:
        st.write("Insight:", text)

    # ---------------- FINAL RECOMMENDATIONS ----------------

    st.header("Strategic Recommendations")

    for col in numeric_cols:

        avg = df[col].mean()
        mx = df[col].max()
        std = df[col].std()

        if avg < mx * 0.6:
            final_recommendations.append(
                f"{col} average is much lower than its maximum value suggesting "
                f"there is room for improvement through targeted actions."
            )

        if std > avg * 0.3:
            final_recommendations.append(
                f"{col} shows large variation. Standardizing processes or "
                f"improving consistency could stabilize performance."
            )

    for rec in final_recommendations:
        st.write("Recommendation:", rec)

    # ---------------- PPT GENERATION ----------------

    if st.button("Generate PPT Report"):

        prs = Presentation()

        slide = prs.slides.add_slide(prs.slide_layouts[0])
        slide.shapes.title.text = "Automated Data Analysis Report"
        slide.placeholders[1].text = "Generated from Excel Dataset"

        for item in reports:

            slide = prs.slides.add_slide(prs.slide_layouts[5])
            slide.shapes.title.text = item["title"]

            tmp1 = tempfile.NamedTemporaryFile(delete=False, suffix=".png")
            item["chart1"].savefig(tmp1.name)

            slide.shapes.add_picture(
                tmp1.name,
                Inches(0.5),
                Inches(1.5),
                height=Inches(3)
            )

            if item["chart2"]:

                tmp2 = tempfile.NamedTemporaryFile(delete=False, suffix=".png")
                item["chart2"].savefig(tmp2.name)

                slide.shapes.add_picture(
                    tmp2.name,
                    Inches(5),
                    Inches(1.5),
                    height=Inches(3)
                )

            textbox = slide.shapes.add_textbox(
                Inches(0.5),
                Inches(4.8),
                Inches(9),
                Inches(1)
            )

            tf = textbox.text_frame
            tf.text = "Insight: " + item["insight"]

        # overall insights slide
        slide = prs.slides.add_slide(prs.slide_layouts[1])
        slide.shapes.title.text = "Overall Observations"

        tf = slide.shapes.placeholders[1].text_frame
        for ins in overall_insights:
            tf.add_paragraph().text = ins

        # recommendation slide
        slide = prs.slides.add_slide(prs.slide_layouts[1])
        slide.shapes.title.text = "Strategic Recommendations"

        tf = slide.shapes.placeholders[1].text_frame
        for rec in final_recommendations:
            tf.add_paragraph().text = rec

        prs.save("analysis_report.pptx")

        with open("analysis_report.pptx","rb") as f:

            st.download_button(
                "Download PPT",
                f,
                file_name="analysis_report.pptx"
            )
