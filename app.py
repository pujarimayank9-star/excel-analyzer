import streamlit as st
import pandas as pd
import matplotlib.pyplot as plt
from pptx import Presentation
from pptx.util import Inches
import tempfile

st.set_page_config(page_title="AI Data Analyst", layout="wide")

st.title("AI Data Analyst & Auto Presentation Generator")

file = st.file_uploader("Upload Excel / CSV", type=["xlsx","csv"])

if file:

    # READ FILE
    if file.name.endswith(".csv"):
        df = pd.read_csv(file)
    else:
        df = pd.read_excel(file)

    st.header("Dataset Preview")
    st.dataframe(df)

    numeric_cols = df.select_dtypes(include=["int64","float64"]).columns
    cat_cols = df.select_dtypes(include=["object"]).columns

    if len(numeric_cols) == 0:
        st.error("No numeric column found")
        st.stop()

    target = numeric_cols[0]

    # ----------------------
    # BASIC STATS
    # ----------------------

    avg = round(df[target].mean(),2)
    mx = df[target].max()
    mn = df[target].min()

    st.header("Key Statistics")

    c1,c2,c3 = st.columns(3)

    c1.metric("Average",avg)
    c2.metric("Max",mx)
    c3.metric("Min",mn)

    # ----------------------
    # TREND ANALYSIS
    # ----------------------

    date_cols = df.select_dtypes(include=["datetime64","object"]).columns

    trend_chart = None

    for col in date_cols:
        try:
            df[col] = pd.to_datetime(df[col])
            trend = df.groupby(col)[target].sum()

            fig,ax = plt.subplots()
            trend.plot(ax=ax)

            st.header("Trend Analysis")
            st.pyplot(fig)

            trend_chart = fig
            break

        except:
            continue

    # ----------------------
    # CATEGORY ANALYSIS
    # ----------------------

    charts = []
    insights = []
    recommendations = []

    for col in cat_cols:

        if df[col].nunique() < 20:

            grp = df.groupby(col)[target].mean()

            best = grp.idxmax()
            worst = grp.idxmin()

            best_val = round(grp.max(),2)
            worst_val = round(grp.min(),2)

            # BAR
            fig_bar,ax = plt.subplots()
            grp.plot(kind="bar",ax=ax)
            ax.set_title(f"{target} by {col}")
            st.pyplot(fig_bar)

            # PIE
            fig_pie,ax = plt.subplots()
            grp.plot(kind="pie",autopct="%1.1f%%",ax=ax)
            ax.set_ylabel("")
            ax.set_title(f"{col} Distribution")
            st.pyplot(fig_pie)

            charts.append((col,fig_bar,fig_pie))

            insight = f"""
{best} performs best in {col} with an average value of {best_val}.
This indicates stronger operational performance or customer demand.

{worst} records the lowest value of {worst_val}.
This suggests weaker performance compared to peers.
"""

            insights.append((col,insight))

            recommendations.append(
                f"Improve {worst} performance in {col} by adopting strategies from {best}."
            )

    # ----------------------
    # CORRELATION
    # ----------------------

    if len(numeric_cols) > 1:

        st.header("Correlation Analysis")

        corr = df[numeric_cols].corr()

        fig,ax = plt.subplots()
        cax = ax.matshow(corr)

        st.pyplot(fig)

    # ----------------------
    # POWERPOINT GENERATION
    # ----------------------

    if st.button("Generate Professional PPT Report"):

        prs = Presentation()

        # TITLE
        slide = prs.slides.add_slide(prs.slide_layouts[0])
        slide.shapes.title.text = "Sales Data Analysis"
        slide.placeholders[1].text = "Automated Business Intelligence Report"

        # STATS
        slide = prs.slides.add_slide(prs.slide_layouts[1])
        slide.shapes.title.text = "Key Statistics"

        slide.placeholders[1].text = f"""
Average Weekly Sales : {avg}

Maximum Weekly Sales : {mx}

Minimum Weekly Sales : {mn}
"""

        # TREND
        if trend_chart:

            tmp = tempfile.NamedTemporaryFile(delete=False,suffix=".png")
            trend_chart.savefig(tmp.name)

            slide = prs.slides.add_slide(prs.slide_layouts[5])
            slide.shapes.title.text = "Trend Analysis"

            slide.shapes.add_picture(
                tmp.name,
                Inches(1),
                Inches(1.5),
                height=Inches(4.5)
            )

        # CATEGORY CHARTS
        for col,bar,pie in charts:

            tmp1 = tempfile.NamedTemporaryFile(delete=False,suffix=".png")
            bar.savefig(tmp1.name)

            tmp2 = tempfile.NamedTemporaryFile(delete=False,suffix=".png")
            pie.savefig(tmp2.name)

            slide = prs.slides.add_slide(prs.slide_layouts[5])
            slide.shapes.title.text = f"{target} by {col}"

            slide.shapes.add_picture(tmp1.name,Inches(0.5),Inches(1.5),height=Inches(4))
            slide.shapes.add_picture(tmp2.name,Inches(6),Inches(1.5),height=Inches(4))

        # INSIGHTS
        for col,text in insights:

            slide = prs.slides.add_slide(prs.slide_layouts[1])
            slide.shapes.title.text = f"{col} Insights"
            slide.placeholders[1].text = text

        # RECOMMENDATIONS
        slide = prs.slides.add_slide(prs.slide_layouts[1])
        slide.shapes.title.text = "Strategic Recommendations"

        slide.placeholders[1].text = "\n".join(recommendations)

        ppt_path = "analysis_report.pptx"

        prs.save(ppt_path)

        with open(ppt_path,"rb") as f:

            st.download_button(
                "Download PPT Report",
                f,
                file_name="business_analysis_report.pptx"
            )
