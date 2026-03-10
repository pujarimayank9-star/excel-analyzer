import streamlit as st
import pandas as pd
import matplotlib.pyplot as plt
from pptx import Presentation
from pptx.util import Inches

st.set_page_config(page_title="Excel Analyzer Pro", layout="wide")

st.title("📊 Excel Smart Data Analyzer")

file = st.file_uploader("Upload Excel File", type=["xlsx"])


def generate_ppt(df):

    prs = Presentation()

    slide = prs.slides.add_slide(prs.slide_layouts[0])
    slide.shapes.title.text = "Excel Data Analysis Report"
    slide.placeholders[1].text = "Automatically generated"

    slide = prs.slides.add_slide(prs.slide_layouts[1])
    slide.shapes.title.text = "Dataset Overview"

    slide.placeholders[1].text = f"""
Rows: {df.shape[0]}
Columns: {df.shape[1]}
"""

    charts = []

    numeric_cols = df.select_dtypes(include="number").columns
    cat_cols = df.select_dtypes(include="object").columns

    # Store comparison
    if "Store" in df.columns and "Weekly_Sales" in df.columns:

        data = df.groupby("Store")["Weekly_Sales"].sum()

        plt.figure()
        data.plot(kind="bar")
        plt.title("Sales by Store")

        file = "store_chart.png"
        plt.savefig(file)
        charts.append(file)

    # Region comparison
    if "Region" in df.columns and "Weekly_Sales" in df.columns:

        data = df.groupby("Region")["Weekly_Sales"].sum()

        plt.figure()
        data.plot(kind="bar")
        plt.title("Sales by Region")

        file = "region_chart.png"
        plt.savefig(file)
        charts.append(file)

    # Distribution chart
    if "Weekly_Sales" in df.columns:

        plt.figure()
        df["Weekly_Sales"].plot(kind="hist")
        plt.title("Sales Distribution")

        file = "distribution.png"
        plt.savefig(file)
        charts.append(file)

    # add charts to ppt
    for c in charts:

        slide = prs.slides.add_slide(prs.slide_layouts[5])
        slide.shapes.title.text = "Chart"
        slide.shapes.add_picture(c, Inches(1), Inches(2), width=Inches(6))

    prs.save("report.pptx")


if file:

    df = pd.read_excel(file)

    st.subheader("Dataset Preview")
    st.dataframe(df)

    # detect columns
    numeric_cols = df.select_dtypes(include="number").columns

    st.subheader("📈 Numeric Metrics")

    if "Weekly_Sales" in df.columns:

        c1,c2,c3 = st.columns(3)

        c1.metric("Average Sales", round(df["Weekly_Sales"].mean(),2))
        c2.metric("Max Sales", df["Weekly_Sales"].max())
        c3.metric("Min Sales", df["Weekly_Sales"].min())

    # Trend chart
    if "Date" in df.columns and "Weekly_Sales" in df.columns:

        st.subheader("📈 Sales Trend")

        df["Date"] = pd.to_datetime(df["Date"])

        trend = df.sort_values("Date")

        fig, ax = plt.subplots()
        ax.plot(trend["Date"], trend["Weekly_Sales"])
        st.pyplot(fig)

    # Store comparison
    if "Store" in df.columns and "Weekly_Sales" in df.columns:

        st.subheader("🏪 Sales by Store")

        store_sales = df.groupby("Store")["Weekly_Sales"].sum()

        fig, ax = plt.subplots()
        store_sales.plot(kind="bar", ax=ax)
        st.pyplot(fig)

        st.success(f"Best Store: {store_sales.idxmax()}")
        st.warning(f"Lowest Store: {store_sales.idxmin()}")

    # Region comparison
    if "Region" in df.columns and "Weekly_Sales" in df.columns:

        st.subheader("🌍 Sales by Region")

        region_sales = df.groupby("Region")["Weekly_Sales"].sum()

        fig, ax = plt.subplots()
        region_sales.plot(kind="bar", ax=ax)
        st.pyplot(fig)

        st.success(f"Best Region: {region_sales.idxmax()}")
        st.warning(f"Lowest Region: {region_sales.idxmin()}")

    # Distribution
    if "Weekly_Sales" in df.columns:

        st.subheader("📊 Sales Distribution")

        fig, ax = plt.subplots()
        df["Weekly_Sales"].plot(kind="hist", ax=ax)
        st.pyplot(fig)

    # Correlation
    if len(numeric_cols) > 1:

        st.subheader("🔗 Correlation Heatmap")

        corr = df[numeric_cols].corr()

        st.dataframe(corr)

        fig, ax = plt.subplots()
        cax = ax.matshow(corr)
        fig.colorbar(cax)

        ax.set_xticks(range(len(corr.columns)))
        ax.set_xticklabels(corr.columns, rotation=90)

        ax.set_yticks(range(len(corr.columns)))
        ax.set_yticklabels(corr.columns)

        st.pyplot(fig)

    st.subheader("🧠 Automatic Insights")

    if "Weekly_Sales" in df.columns:

        avg = df["Weekly_Sales"].mean()

        if avg > 2000:
            st.write("Sales performance is strong.")
        elif avg > 1200:
            st.write("Sales performance is moderate.")
        else:
            st.write("Sales performance is weak.")

    if st.button("📥 Generate PowerPoint Report"):

        generate_ppt(df)

        with open("report.pptx","rb") as f:

            st.download_button(
                "Download PPT Report",
                f,
                file_name="analysis_report.pptx"
            )
