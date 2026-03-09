import streamlit as st
import pandas as pd
import matplotlib.pyplot as plt

st.set_page_config(page_title="Smart Excel Analyzer", page_icon="📊", layout="wide")

st.title("📊 Smart Excel Analytics Dashboard")
st.write("Upload ANY Excel file and get automatic insights")

file = st.file_uploader("Upload Excel File", type=["xlsx"])

if file is not None:

    df = pd.read_excel(file)

    st.subheader("Raw Data")
    st.dataframe(df)

    # ==========================
    # Detect columns
    # ==========================

    numeric_cols = df.select_dtypes(include=["number"]).columns
    cat_cols = df.select_dtypes(include=["object"]).columns

    # ==========================
    # Automatic Charts
    # ==========================

    st.header("📊 Automatic Charts")

    for col in numeric_cols:

        fig, ax = plt.subplots()

        df[col].plot(kind="hist", ax=ax)

        ax.set_title(f"{col} Distribution")

        st.pyplot(fig)

    # ==========================
    # Category vs Number
    # ==========================

    st.header("📊 Category vs Number Comparisons")

    for cat in cat_cols:
        for num in numeric_cols:

            try:

                data = df.groupby(cat)[num].mean()

                fig, ax = plt.subplots()

                data.plot(kind="bar", ax=ax)

                ax.set_title(f"{num} by {cat}")

                st.pyplot(fig)

                best = data.idxmax()
                worst = data.idxmin()

                st.write(f"🏆 Best {cat}: {best}")
                st.write(f"⚠ Lowest {cat}: {worst}")

            except:
                pass

    # ==========================
    # Correlation
    # ==========================

    if len(numeric_cols) > 1:

        st.header("🔎 Correlation Analysis")

        corr = df[numeric_cols].corr()

        st.dataframe(corr)

        fig, ax = plt.subplots()

        cax = ax.matshow(corr)

        plt.xticks(range(len(corr.columns)), corr.columns, rotation=90)
        plt.yticks(range(len(corr.columns)), corr.columns)

        fig.colorbar(cax)

        st.pyplot(fig)

    # ==========================
    # Top Performers
    # ==========================

    st.header("🏆 Top Performers")

    for cat in cat_cols:
        for num in numeric_cols:

            try:

                top = df.groupby(cat)[num].sum().sort_values(ascending=False).head(5)

                st.write(f"Top {cat} by {num}")

                st.dataframe(top)

            except:
                pass

    # ==========================
    # Profit Analysis
    # ==========================

    if "Profit" in df.columns and "Product_Category" in df.columns:

        st.header("💰 Profit Analysis")

        profit = df.groupby("Product_Category")["Profit"].sum()

        fig, ax = plt.subplots()

        profit.plot(kind="bar", ax=ax)

        st.pyplot(fig)

        st.write(f"🏆 Most Profitable Product: {profit.idxmax()}")
        st.write(f"⚠ Least Profitable Product: {profit.idxmin()}")

    # ==========================
    # Sales Trend (SAFE VERSION)
    # ==========================

    if "Date" in df.columns and "Weekly_Sales" in df.columns:

        st.header("📈 Sales Trend")

        try:

            df["Date"] = pd.to_datetime(df["Date"], errors="coerce")

            df2 = df.dropna(subset=["Date"])

            trend = df2.groupby("Date")["Weekly_Sales"].sum()

            fig, ax = plt.subplots()

            trend.plot(ax=ax)

            ax.set_title("Sales Over Time")

            st.pyplot(fig)

        except:

            st.warning("Date format not supported for trend analysis.")

    # ==========================
    # Smart Insights
    # ==========================

    st.header("🧠 Smart Insights")

    for cat in cat_cols:
        for num in numeric_cols:

            try:

                data = df.groupby(cat)[num].sum()

                best = data.idxmax()
                worst = data.idxmin()

                st.write(f"Highest {num} comes from {best} in {cat}")
                st.write(f"Lowest {num} comes from {worst} in {cat}")

            except:
                pass
