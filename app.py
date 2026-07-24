import streamlit as st
import pandas as pd
import plotly.express as px
import plotly.graph_objects as go

from analysis import *
from charts import *
from ai_engine import *
from ppt_generator import *
from pdf_generator import *

st.set_page_config(
    page_title="AI Smart Excel Analyst",
    page_icon="📊",
    layout="wide"
)

st.title("📊 AI Smart Excel Data Analyst")
st.caption("Powered by Gemini AI")

uploaded_file = st.file_uploader(
    "Upload Excel or CSV",
    type=["xlsx","csv"]
)

if uploaded_file:

    if uploaded_file.name.endswith(".csv"):
        df = pd.read_csv(uploaded_file)

    else:
        df = pd.read_excel(uploaded_file)

    st.success("Dataset Loaded Successfully")

    st.write(df.head())

    dataset_overview(df)

    dataset_detection(df)

    missing_analysis(df)

    duplicate_analysis(df)

    numeric_analysis(df)

    category_analysis(df)

    correlation_analysis(df)

    outlier_analysis(df)

    trend_analysis(df)

    ai_business_analysis(df)

    ppt_download(df)

    pdf_download(df)
