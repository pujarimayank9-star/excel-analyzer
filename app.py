import streamlit as st
import pandas as pd
import numpy as np
import matplotlib.pyplot as plt
from pptx import Presentation

st.set_page_config(page_title="AI Business Analyst", layout="wide")

st.title("AI Business Analyst")

file = st.file_uploader("Upload Excel or CSV", type=["xlsx","csv"])

def generate_ppt(insights,suggestions):

    prs = Presentation()

    slide = prs.slides.add_slide(prs.slide_layouts[0])
    slide.shapes.title.text="Business Analysis Report"
    slide.placeholders[1].text="Generated automatically"

    slide = prs.slides.add_slide(prs.slide_layouts[1])
    slide.shapes.title.text="Insights"
    slide.placeholders[1].text="\n".join(insights)

    slide = prs.slides.add_slide(prs.slide_layouts[1])
    slide.shapes.title.text="Recommendations"
    slide.placeholders[1].text="\n".join(suggestions)

    prs.save("report.pptx")


if file:

    if file.name.endswith(".csv"):
        df=pd.read_csv(file)
    else:
        df=pd.read_excel(file)

    st.header("Dataset Preview")
    st.dataframe(df)

# -----------------------
# DATE DETECTION FIRST
# -----------------------

    date_cols=[]

    for col in df.columns:

        parsed=pd.to_datetime(df[col],dayfirst=True,errors="coerce")

        if parsed.notna().sum()>len(df)*0.6:
            df[col]=parsed
            date_cols.append(col)

# -----------------------
# NUMERIC DETECTION
# -----------------------

    numeric_cols=[]

    for col in df.columns:

        if col in date_cols:
            continue

        converted=pd.to_numeric(df[col],errors="coerce")

        if converted.notna().sum()>len(df)*0.6:
            df[col]=converted
            numeric_cols.append(col)

    cat_cols=[c for c in df.columns if c not in numeric_cols and c not in date_cols]

    insights=[]
    suggestions=[]

# -----------------------
# BASIC STATS
# -----------------------

    st.header("Basic Statistics")

    for col in numeric_cols:

        st.subheader(col)

        avg=df[col].mean()
        mx=df[col].max()
        mn=df[col].min()

        c1,c2,c3=st.columns(3)

        c1.metric("Average",round(avg,2))
        c2.metric("Max",mx)
        c3.metric("Min",mn)

        fig,ax=plt.subplots()

        df[col].plot(kind="hist",ax=ax)

        st.pyplot(fig)

        insights.append(f"{col} average value is {round(avg,2)}")

# -----------------------
# CATEGORY ANALYSIS
# -----------------------

    st.header("Category Analysis")

    for cat in cat_cols:

        if df[cat].nunique()<20:

            for num in numeric_cols:

                grp=df.groupby(cat)[num].mean()

                fig,ax=plt.subplots()

                grp.plot(kind="bar",ax=ax)

                st.pyplot(fig)

                best=grp.idxmax()
                worst=grp.idxmin()

                insights.append(f"{best} performs best in {cat}")
                insights.append(f"{worst} performs weakest in {cat}")

# -----------------------
# PIE CHARTS
# -----------------------

    st.header("Distribution")

    for cat in cat_cols:

        if df[cat].nunique()<10:

            fig,ax=plt.subplots()

            df[cat].value_counts().plot(kind="pie",autopct='%1.1f%%',ax=ax)

            ax.set_ylabel("")

            st.subheader(cat)

            st.pyplot(fig)

# -----------------------
# TREND ANALYSIS
# -----------------------

    st.header("Trend Analysis")

    if date_cols:

        date_col=date_cols[0]

        st.write("Detected date column:",date_col)

        df=df.sort_values(date_col)

        for num in numeric_cols:

            daily=df.groupby(date_col)[num].sum()

            fig,ax=plt.subplots()

            daily.plot(ax=ax)

            ax.set_title("Daily Trend")

            st.pyplot(fig)

            df["month"]=df[date_col].dt.to_period("M")

            monthly=df.groupby("month")[num].sum()

            fig,ax=plt.subplots()

            monthly.plot(ax=ax)

            ax.set_title("Monthly Trend")

            st.pyplot(fig)

            ma=daily.rolling(7).mean()

            fig,ax=plt.subplots()

            daily.plot(ax=ax,label="Actual")

            ma.plot(ax=ax,label="7 Day Moving Avg")

            ax.legend()

            st.pyplot(fig)

            df["weekday"]=df[date_col].dt.day_name()

            season=df.groupby("weekday")[num].mean()

            fig,ax=plt.subplots()

            season.plot(kind="bar",ax=ax)

            ax.set_title("Seasonality")

            st.pyplot(fig)

            best_day=season.idxmax()

            insights.append(f"Highest sales generally occur on {best_day}")

    else:

        st.info("No date column detected")

# -----------------------
# SUGGESTIONS
# -----------------------

    if "Region" in df.columns and numeric_cols:

        region_sales=df.groupby("Region")[numeric_cols[0]].sum()

        best_region=region_sales.idxmax()

        suggestions.append(
        f"{best_region} region generates strongest revenue. Expanding product availability in similar regions may increase sales."
        )

    if "Product_Category" in df.columns and numeric_cols:

        cat_sales=df.groupby("Product_Category")[numeric_cols[0]].sum()

        weak=cat_sales.idxmin()

        suggestions.append(
        f"{weak} category contributes lowest revenue. Targeted promotions or pricing adjustments could improve performance."
        )

# -----------------------
# INSIGHTS
# -----------------------

    st.header("Business Insights")

    for i in insights:
        st.write("•",i)

# -----------------------
# RECOMMENDATIONS
# -----------------------

    st.header("Strategic Suggestions")

    for s in suggestions:
        st.write("•",s)

# -----------------------
# PPT
# -----------------------

    if st.button("Generate PowerPoint Report"):

        generate_ppt(insights,suggestions)

        with open("report.pptx","rb") as f:

            st.download_button("Download PPT",f,file_name="analysis_report.pptx")
