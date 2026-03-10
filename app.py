# app.py

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

    # -------- LOAD DATA --------
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

    # -------- NUMERIC ANALYSIS --------

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

    # -------- CATEGORY ANALYSIS --------

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

    # -------- CORRELATION ANALYSIS --------

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

        correlation_explanation = ""

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

                correlation_explanation += f"{col1} vs {col2} shows {relation} relationship ({round(val,2)}). "

                st.write("Insight:", f"{col1} vs {col2} → correlation {round(val,2)} ({relation}).")

        reports.append({
            "title": "Correlation",
            "chart1": fig,
            "chart2": None,
            "insight": correlation_explanation
        })

    # -------- DATASET LEVEL OBSERVATION --------

    st.header("Dataset Level Observations")

    observations = []

    if len(numeric_cols) > 0:

        variability = df[numeric_cols].std().mean()
        avg_value = df[numeric_cols].mean().mean()

        if variability > avg_value*0.3:
            observations.append(
                "Numeric metrics vary significantly across the dataset indicating inconsistent performance."
            )
        else:
            observations.append(
                "Numeric values appear relatively stable across observations."
            )

    if len(numeric_cols) > 1:

        corr_values = df[numeric_cols].corr().abs().values.flatten()
        corr_values = corr_values[corr_values < 1]

        if len(corr_values)>0:

            strongest = corr_values.max()

            if strongest > 0.6:
                observations.append(
                    "Some variables show strong interaction meaning improvement in one may influence others."
                )
            else:
                observations.append(
                    "Most variables behave independently with limited interaction."
                )

    for obs in observations:
        st.write("Observation:", obs)

    # -------- STRATEGIC RECOMMENDATIONS --------

    st.header("Strategic Recommendations")

    recommendations = []

    if len(numeric_cols)>0:

        avg_vals = df[numeric_cols].mean().sort_values()

        weakest = avg_vals.index[0]
        strongest = avg_vals.index[-1]

        recommendations.append(
            f"{weakest} appears weakest on average while {strongest} performs strongest."
        )

        recommendations.append(
            "Focus improvement strategies on weaker metrics while replicating practices used in stronger ones."
        )

    if len(cat_cols)>0:

        recommendations.append(
            "Review categorical distribution to ensure balanced representation across categories."
        )

    for r in recommendations:
        st.write("Recommendation:", r)

    # -------- PPT GENERATION --------

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
