# app.py

import streamlit as st
import pandas as pd
import matplotlib.pyplot as plt
from pptx import Presentation
from pptx.util import Inches
import tempfile

st.set_page_config(layout="wide")
st.title("AI Smart Excel Data Analyst")

file = st.file_uploader("Upload Excel / CSV", type=["xlsx","csv"])

if file:

    # ---------- LOAD DATA ----------
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

    # ---------- DATASET TYPE DETECTION ----------

    dataset_type = "generic"

    columns_lower = [c.lower() for c in df.columns]

    if any(x in columns_lower for x in ["sales","revenue","profit"]):
        dataset_type = "sales"

    if any(x in columns_lower for x in ["marks","score","grade","attendance"]):
        dataset_type = "education"

    st.write("Detected Dataset Type:", dataset_type)

    # ---------- NUMERIC ANALYSIS ----------

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

        insight = f"{col} ranges from {mn} to {mx} with an average of {round(avg,2)}."

        st.write("Insight:", insight)

        reports.append({
            "title": col,
            "chart1": fig,
            "chart2": None,
            "insight": insight
        })

    # ---------- CATEGORY ANALYSIS ----------

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

            insight = f"{best} represents the largest share in {cat} contributing about {percent}%."

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

    # ---------- CORRELATION ANALYSIS ----------

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
                        ha="center", va="center")

        st.pyplot(fig)

        correlation_text = ""

        for i in range(len(corr.columns)):
            for j in range(i+1, len(corr.columns)):

                col1 = corr.columns[i]
                col2 = corr.columns[j]
                val = corr.iloc[i,j]

                if abs(val) > 0.6:
                    relation = "strong"
                elif abs(val) > 0.3:
                    relation = "moderate"
                else:
                    relation = "weak"

                text = f"{col1} vs {col2} shows {relation} relationship ({round(val,2)})."

                st.write("Insight:", text)

                correlation_text += text + " "

        reports.append({
            "title": "Correlation",
            "chart1": fig,
            "chart2": None,
            "insight": correlation_text
        })

    # ---------- SMART DATASET OBSERVATION ----------

    st.header("Dataset Observations")

    observations = []

    if dataset_type == "sales":

        observations.append(
            "Sales datasets typically depend on regional performance, seasonality and product demand."
        )

        if len(numeric_cols) > 0:

            variability = df[numeric_cols].std().mean()

            if variability > df[numeric_cols].mean().mean()*0.3:

                observations.append(
                    "Sales values vary significantly which may indicate inconsistent demand across regions or time periods."
                )

            else:

                observations.append(
                    "Sales appear relatively stable suggesting consistent demand patterns."
                )

    elif dataset_type == "education":

        observations.append(
            "Academic datasets often reveal performance differences across subjects or attendance patterns."
        )

        if len(numeric_cols) > 0:

            variability = df[numeric_cols].std().mean()

            if variability > df[numeric_cols].mean().mean()*0.3:

                observations.append(
                    "Student performance varies across subjects indicating uneven academic strengths."
                )

            else:

                observations.append(
                    "Student performance appears relatively consistent across metrics."
                )

    else:

        observations.append(
            "The dataset contains general numeric and categorical variables which may represent operational or survey data."
        )

    for obs in observations:
        st.write("Observation:", obs)

    # ---------- SMART RECOMMENDATIONS ----------

    st.header("Strategic Recommendations")

    recommendations = []

    if dataset_type == "sales":

        recommendations.append(
            "Identify high performing segments and replicate successful strategies across weaker regions or periods."
        )

        recommendations.append(
            "Analyze time based patterns to detect seasonal demand and adjust inventory or marketing accordingly."
        )

    elif dataset_type == "education":

        recommendations.append(
            "Provide targeted academic support for subjects where performance variability is highest."
        )

        recommendations.append(
            "Use attendance and performance patterns to identify students needing additional assistance."
        )

    else:

        recommendations.append(
            "Investigate key drivers behind high and low values to understand operational performance."
        )

        recommendations.append(
            "Focus improvement strategies on weaker performing metrics while maintaining strengths."
        )

    for r in recommendations:
        st.write("Recommendation:", r)

    # ---------- PPT GENERATION ----------

    if st.button("Generate PPT Report"):

        prs = Presentation()

        slide = prs.slides.add_slide(prs.slide_layouts[0])
        slide.shapes.title.text = "AI Data Analysis Report"
        slide.placeholders[1].text = "Generated automatically"

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

            textbox.text = "Insight: " + item["insight"]

        slide = prs.slides.add_slide(prs.slide_layouts[1])
        slide.shapes.title.text = "Dataset Observations"

        tf = slide.shapes.placeholders[1].text_frame

        for obs in observations:
            tf.add_paragraph().text = obs

        slide = prs.slides.add_slide(prs.slide_layouts[1])
        slide.shapes.title.text = "Strategic Recommendations"

        tf = slide.shapes.placeholders[1].text_frame

        for r in recommendations:
            tf.add_paragraph().text = r

        prs.save("analysis_report.pptx")

        with open("analysis_report.pptx","rb") as f:

            st.download_button(
                "Download PPT",
                f,
                file_name="analysis_report.pptx"
            )
