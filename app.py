import streamlit as st
import pandas as pd
import numpy as np
import matplotlib.pyplot as plt
from pptx import Presentation

st.set_page_config(page_title="AI Data Analyst", layout="wide")

st.title("📊 Universal AI Data Analyst")

uploaded_file = st.file_uploader("Upload Excel File", type=["xlsx","csv"])


def generate_ppt(insights):

    prs = Presentation()

    slide = prs.slides.add_slide(prs.slide_layouts[0])
    slide.shapes.title.text = "AI Data Analysis Report"
    slide.placeholders[1].text = "Automatically generated insights"

    slide = prs.slides.add_slide(prs.slide_layouts[1])
    slide.shapes.title.text = "Key Insights"
    slide.placeholders[1].text = "\n".join(insights)

    prs.save("analysis_report.pptx")


if uploaded_file:

    if uploaded_file.name.endswith(".csv"):
        df = pd.read_csv(uploaded_file)
    else:
        df = pd.read_excel(uploaded_file)

    st.subheader("Dataset Preview")
    st.dataframe(df)

    numeric_cols = df.select_dtypes(include=np.number).columns.tolist()
    cat_cols = df.select_dtypes(include="object").columns.tolist()

    insights = []

# -------------------------
# DATASET OVERVIEW
# -------------------------

    st.header("Dataset Overview")

    c1,c2,c3 = st.columns(3)

    c1.metric("Rows", df.shape[0])
    c2.metric("Columns", df.shape[1])
    c3.metric("Numeric Columns", len(numeric_cols))

# -------------------------
# NUMERIC ANALYSIS
# -------------------------

    if numeric_cols:

        st.header("Numeric Analysis")

        for col in numeric_cols:

            avg = df[col].mean()
            mx = df[col].max()
            mn = df[col].min()

            c1,c2,c3 = st.columns(3)

            c1.metric(f"{col} Average", round(avg,2))
            c2.metric(f"{col} Max", mx)
            c3.metric(f"{col} Min", mn)

            insights.append(f"{col} average value is {round(avg,2)}")

            fig, ax = plt.subplots()

            df[col].plot(kind="hist", ax=ax)

            ax.set_title(f"{col} Distribution")

            st.pyplot(fig)

            st.write(f"{col} distribution shows spread of values.")

# -------------------------
# CATEGORY BASED ANALYSIS
# -------------------------

    if cat_cols and numeric_cols:

        st.header("Category Based Analysis")

        for cat in cat_cols:

            for num in numeric_cols:

                try:

                    grouped = df.groupby(cat)[num].mean()

                    fig, ax = plt.subplots()

                    grouped.plot(kind="bar", ax=ax)

                    ax.set_title(f"{cat} vs {num}")

                    st.pyplot(fig)

                    best = grouped.idxmax()
                    worst = grouped.idxmin()

                    st.write(f"Best {cat} based on {num}: **{best}**")
                    st.write(f"Worst {cat} based on {num}: **{worst}**")

                    insights.append(f"{best} performs best in {cat} for {num}")
                    insights.append(f"{worst} performs worst in {cat} for {num}")

                except:
                    pass

# -------------------------
# CATEGORY CONTRIBUTION
# -------------------------

    if cat_cols:

        st.header("Category Contribution")

        for cat in cat_cols:

            try:

                counts = df[cat].value_counts()

                fig, ax = plt.subplots()

                counts.plot(kind="pie", autopct="%1.1f%%", ax=ax)

                ax.set_ylabel("")

                st.pyplot(fig)

                insights.append(f"{cat} distribution analyzed")

            except:
                pass

# -------------------------
# CORRELATION
# -------------------------

    if len(numeric_cols) > 1:

        st.header("Correlation Heatmap")

        try:

            corr = df[numeric_cols].corr()

            fig, ax = plt.subplots()

            cax = ax.matshow(corr)

            fig.colorbar(cax)

            ax.set_xticks(range(len(corr.columns)))
            ax.set_xticklabels(corr.columns, rotation=90)

            ax.set_yticks(range(len(corr.columns)))
            ax.set_yticklabels(corr.columns)

            st.pyplot(fig)

            insights.append("Correlation between numeric variables analyzed")

        except:
            pass

# -------------------------
# NUMERIC RELATIONSHIPS
# -------------------------

    if len(numeric_cols) > 1:

        st.header("Numeric Relationships")

        for i in range(len(numeric_cols)-1):

            try:

                x = numeric_cols[i]
                y = numeric_cols[i+1]

                fig, ax = plt.subplots()

                ax.scatter(df[x], df[y])

                ax.set_xlabel(x)
                ax.set_ylabel(y)

                ax.set_title(f"{x} vs {y}")

                st.pyplot(fig)

            except:
                pass

# -------------------------
# TREND ANALYSIS (SAFE)
# -------------------------

    st.header("Trend Analysis")

    date_column = None

    for col in df.columns:

        try:

            converted = pd.to_datetime(df[col], errors="coerce")

            if converted.notna().sum() > len(df)*0.8:

                date_column = col
                df[col] = converted
                break

        except:
            pass

    if date_column:

        st.write(f"Detected date column: {date_column}")

        df = df.dropna(subset=[date_column])

        for num in numeric_cols:

            try:

                trend = df[[date_column,num]].dropna()

                trend = trend.sort_values(date_column)

                fig, ax = plt.subplots()

                ax.plot(trend[date_column], trend[num])

                ax.set_title(f"{num} Trend Over Time")

                st.pyplot(fig)

            except:
                pass

    else:

        st.info("No proper date column detected. Trend analysis skipped.")

# -------------------------
# AI INSIGHTS
# -------------------------

    st.header("AI Insights")

    for i in insights:

        st.write("•", i)

# -------------------------
# PPT GENERATOR
# -------------------------

    if st.button("Generate PowerPoint Report"):

        generate_ppt(insights)

        with open("analysis_report.pptx","rb") as f:

            st.download_button(
                "Download PPT",
                f,
                file_name="AI_Data_Analysis_Report.pptx"
            )
