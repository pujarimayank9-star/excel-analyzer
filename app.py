import streamlit as st
import pandas as pd
import matplotlib.pyplot as plt

st.set_page_config(
    page_title="Smart Excel Analyzer",
    page_icon="📊",
    layout="wide"
)

st.title("📊 Smart Excel Analyzer")
st.write("Upload ANY Excel file and get deep analysis & comparisons")

file = st.file_uploader("Upload Excel File", type=["xlsx","csv"])

if file is not None:

    if file.name.endswith(".csv"):
        df = pd.read_csv(file)
    else:
        df = pd.read_excel(file)

    st.subheader("Raw Data")
    st.dataframe(df)

    numeric_cols = df.select_dtypes(include="number").columns
    cat_cols = df.select_dtypes(include="object").columns

    # ----------------------
    # Numeric statistics
    # ----------------------

    st.subheader("📈 Numeric Analysis")

    for col in numeric_cols:

        avg = df[col].mean()
        max_val = df[col].max()
        min_val = df[col].min()

        c1,c2,c3 = st.columns(3)

        with c1:
            st.metric(f"Average {col}", round(avg,2))

        with c2:
            st.metric(f"Max {col}", max_val)

        with c3:
            st.metric(f"Min {col}", min_val)

        st.bar_chart(df[col])

        st.write(f"Range of {col}: {max_val-min_val}")

    # ----------------------
    # Category comparisons
    # ----------------------

    st.subheader("📊 Category vs Number Comparisons")

    for cat in cat_cols:

        if df[cat].nunique() < 15:

            for num in numeric_cols:

                st.write(f"{num} by {cat}")

                grouped = df.groupby(cat)[num].mean()

                st.bar_chart(grouped)

                top = grouped.idxmax()
                bottom = grouped.idxmin()

                st.write(f"🏆 Best {cat}: {top}")
                st.write(f"⚠ Lowest {cat}: {bottom}")

    # ----------------------
    # Correlation analysis
    # ----------------------

    if len(numeric_cols) > 1:

        st.subheader("🔍 Correlation Analysis")

        corr = df[numeric_cols].corr()

        st.write(corr)

        fig, ax = plt.subplots()
        cax = ax.matshow(corr)
        plt.xticks(range(len(corr.columns)), corr.columns, rotation=90)
        plt.yticks(range(len(corr.columns)), corr.columns)

        st.pyplot(fig)

    # ----------------------
    # Deep insights
    # ----------------------

    st.subheader("🤖 Automatic Insights")

    for num in numeric_cols:

        avg = df[num].mean()

        if avg > df[num].max()*0.7:
            st.write(f"{num} values are generally high.")

        elif avg < df[num].max()*0.4:
            st.write(f"{num} values are generally low.")

        else:
            st.write(f"{num} values are moderately distributed.")

    st.write("Analysis complete.")
