import streamlit as st
import pandas as pd
import matplotlib.pyplot as plt

st.set_page_config(page_title="AI Business Analyzer", page_icon="📊", layout="wide")

st.title("📊 AI Excel Business Analyzer")
st.write("Upload ANY Excel file → Automatic Charts → Insights → Recommendations")

file = st.file_uploader("Upload Excel file", type=["xlsx"])

if file is not None:

    df = pd.read_excel(file)

    st.subheader("📄 Raw Data")
    st.dataframe(df)

    numeric_cols = df.select_dtypes(include=["number"]).columns
    cat_cols = df.select_dtypes(include=["object"]).columns

    insights = []
    recommendations = []

    # ------------------------
    # CATEGORY vs NUMBER
    # ------------------------

    st.header("📊 Category Analysis")

    for cat in cat_cols:
        for num in numeric_cols:

            try:
                data = df.groupby(cat)[num].sum()

                fig, ax = plt.subplots()
                data.plot(kind="bar", ax=ax)
                ax.set_title(f"{num} by {cat}")

                st.pyplot(fig)

                best = data.idxmax()
                worst = data.idxmin()

                st.write(f"🏆 Best {cat}: {best}")
                st.write(f"⚠ Lowest {cat}: {worst}")

                insights.append(
                    f"{best} generates the highest {num} in {cat}, while {worst} performs the lowest."
                )

                recommendations.append(
                    f"Focus more resources on {best} and investigate improvement strategies for {worst}."
                )

            except:
                pass

    # ------------------------
    # CORRELATION
    # ------------------------

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

    # ------------------------
    # TOP PERFORMERS
    # ------------------------

    st.header("🏆 Top Performers")

    for cat in cat_cols:
        for num in numeric_cols:

            try:

                top = df.groupby(cat)[num].sum().sort_values(ascending=False).head(5)

                st.write(f"Top {cat} by {num}")
                st.dataframe(top)

            except:
                pass

    # ------------------------
    # SALES TREND
    # ------------------------

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

            st.warning("Date format not supported.")

    # ------------------------
    # AI INSIGHTS
    # ------------------------

    st.header("🧠 AI Business Insights")

    if len(insights) > 0:

        for i in insights[:10]:
            st.write("•", i)

    else:

        st.write("Not enough categorical data.")

    # ------------------------
    # BUSINESS RECOMMENDATIONS
    # ------------------------

    st.header("🚀 Business Recommendations")

    if len(recommendations) > 0:

        for r in recommendations[:10]:
            st.write("•", r)

    else:

        st.write("No recommendations generated.")
