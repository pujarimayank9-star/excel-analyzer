import streamlit as st
import pandas as pd
import numpy as np
import matplotlib.pyplot as plt
from pptx import Presentation
from pptx.util import Inches

st.set_page_config(page_title="AI Data Analyst", layout="wide")

st.title("📊 Universal AI Data Analyst")

file = st.file_uploader("Upload Excel File", type=["xlsx","csv"])


def generate_ppt(insights):

    prs = Presentation()

    slide = prs.slides.add_slide(prs.slide_layouts[0])
    slide.shapes.title.text = "AI Data Analysis Report"
    slide.placeholders[1].text = "Auto generated insights"

    slide = prs.slides.add_slide(prs.slide_layouts[1])
    slide.shapes.title.text = "Key Insights"
    slide.placeholders[1].text = "\n".join(insights)

    prs.save("report.pptx")


if file:

    if file.name.endswith(".csv"):
        df = pd.read_csv(file)
    else:
        df = pd.read_excel(file)

    st.subheader("Dataset Preview")
    st.dataframe(df)

    numeric_cols = df.select_dtypes(include=np.number).columns
    cat_cols = df.select_dtypes(include="object").columns

    insights = []

    st.subheader("Dataset Overview")

    c1,c2,c3 = st.columns(3)

    c1.metric("Rows", df.shape[0])
    c2.metric("Columns", df.shape[1])
    c3.metric("Numeric Columns", len(numeric_cols))

    # Numeric analysis

    if len(numeric_cols)>0:

        st.subheader("Numeric Analysis")

        for col in numeric_cols:

            avg = df[col].mean()
            mx = df[col].max()
            mn = df[col].min()

            c1,c2,c3 = st.columns(3)

            c1.metric(f"{col} Avg", round(avg,2))
            c2.metric(f"{col} Max", mx)
            c3.metric(f"{col} Min", mn)

            insights.append(f"{col} average value is {round(avg,2)}")

            fig, ax = plt.subplots()

            df[col].plot(kind="hist", ax=ax)

            ax.set_title(f"{col} Distribution")

            st.pyplot(fig)

            st.write(f"{col} distribution shows how values are spread across dataset.")

    # Category vs numeric

    if len(cat_cols)>0 and len(numeric_cols)>0:

        st.subheader("Category Based Analysis")

        for cat in cat_cols:

            for num in numeric_cols:

                data = df.groupby(cat)[num].mean()

                fig, ax = plt.subplots()

                data.plot(kind="bar", ax=ax)

                ax.set_title(f"{cat} vs {num}")

                st.pyplot(fig)

                best = data.idxmax()
                worst = data.idxmin()

                st.write(f"Best {cat} based on {num}: **{best}**")
                st.write(f"Worst {cat} based on {num}: **{worst}**")

                insights.append(f"{best} performs best in {cat} based on {num}")
                insights.append(f"{worst} performs lowest in {cat} based on {num}")

    # Category contribution

    if len(cat_cols)>0:

        st.subheader("Category Contribution")

        for cat in cat_cols:

            counts = df[cat].value_counts()

            fig, ax = plt.subplots()

            counts.plot(kind="pie", ax=ax, autopct="%1.1f%%")

            ax.set_ylabel("")

            st.pyplot(fig)

            insights.append(f"{cat} distribution shows category contribution in dataset")

    # Scatter relationships

    if len(numeric_cols)>1:

        st.subheader("Numeric Relationships")

        for i in range(len(numeric_cols)-1):

            x = numeric_cols[i]
            y = numeric_cols[i+1]

            fig, ax = plt.subplots()

            ax.scatter(df[x], df[y])

            ax.set_xlabel(x)
            ax.set_ylabel(y)

            st.pyplot(fig)

            st.write(f"Relationship between {x} and {y}")

    # Correlation

    if len(numeric_cols)>1:

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

        insights.append("Correlation heatmap shows relationships between numeric variables")

    # Date trend

    for col in df.columns:

        if "date" in col.lower():

            st.subheader("Trend Analysis")

            df[col] = pd.to_datetime(df[col])

            for num in numeric_cols:

                trend = df.groupby(col)[num].sum()

                fig, ax = plt.subplots()

                trend.plot(ax=ax)

                ax.set_title(f"{num} Trend")

                st.pyplot(fig)

                insights.append(f"{num} trend over time analyzed")

    # AI Insights

    st.subheader("AI Insights")

    for i in insights:

        st.write("•", i)

    # PPT

    if st.button("Generate PowerPoint Report"):

        generate_ppt(insights)

        with open("report.pptx","rb") as f:

            st.download_button(
                "Download PPT",
                f,
                file_name="analysis_report.pptx"
            )
