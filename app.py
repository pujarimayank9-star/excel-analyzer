import streamlit as st
import pandas as pd

st.set_page_config(
    page_title="Universal Excel Analyzer",
    page_icon="📊",
    layout="wide"
)

st.title("📊 Universal Excel Analyzer")
st.write("Upload ANY Excel file and get automatic charts, insights and analysis")

file = st.file_uploader("Upload Excel File", type=["xlsx","csv"])

if file is not None:

    # Read file
    if file.name.endswith(".csv"):
        df = pd.read_csv(file)
    else:
        df = pd.read_excel(file)

    st.subheader("Raw Data")
    st.dataframe(df)

    numeric_cols = df.select_dtypes(include="number").columns
    text_cols = df.select_dtypes(include="object").columns

    st.subheader("📊 Automatic Charts")

    # Numeric charts
    for col in numeric_cols:
        st.write(f"{col} Distribution")
        st.bar_chart(df[col])

    # Category charts
    for col in text_cols:
        counts = df[col].value_counts()
        if len(counts) < 20:
            st.write(f"{col} Categories")
            st.bar_chart(counts)

    st.subheader("📈 Automatic Insights")

    for col in numeric_cols:

        avg = round(df[col].mean(),2)
        maximum = df[col].max()
        minimum = df[col].min()

        c1,c2,c3 = st.columns(3)

        with c1:
            st.metric(f"Average {col}",avg)

        with c2:
            st.metric(f"Max {col}",maximum)

        with c3:
            st.metric(f"Min {col}",minimum)

        # AI-style analysis
        st.subheader(f"🤖 Analysis for {col}")

        if avg > (maximum * 0.7):
            st.write(f"The average {col} is relatively high compared to the maximum value.")

        elif avg < (maximum * 0.4):
            st.write(f"The average {col} is relatively low compared to the maximum value.")

        else:
            st.write(f"The {col} values are moderately distributed.")

        st.write(f"Range of {col} is {maximum - minimum}.")
