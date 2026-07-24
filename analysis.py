import streamlit as st
import pandas as pd

def dataset_overview(df):

    st.header("📊 Dataset Dashboard")

    rows = df.shape[0]
    cols = df.shape[1]
    missing = int(df.isnull().sum().sum())
    duplicates = int(df.duplicated().sum())

    c1, c2, c3, c4 = st.columns(4)

    c1.metric("Rows", rows)
    c2.metric("Columns", cols)
    c3.metric("Missing", missing)
    c4.metric("Duplicates", duplicates)

    st.divider()

    with st.expander("📄 Dataset Preview", expanded=True):
        st.dataframe(df.head(10), use_container_width=True)

    st.subheader("Column Types")

    dtype_df = pd.DataFrame({
        "Column": df.columns,
        "Datatype": df.dtypes.astype(str)
    })

    st.dataframe(dtype_df, use_container_width=True)


def dataset_detection(df):
    st.header("🔍 Dataset Type")

    cols = [c.lower() for c in df.columns]

    dtype = "Generic"

    if any(x in cols for x in ["sales", "revenue", "profit"]):
        dtype = "Sales"

    elif any(x in cols for x in ["marks", "grade", "score"]):
        dtype = "Education"

    elif any(x in cols for x in ["employee", "salary"]):
        dtype = "HR"

    elif any(x in cols for x in ["product", "stock", "inventory"]):
        dtype = "Inventory"

    st.success(f"Detected Dataset: {dtype}")


def missing_analysis(df):
    st.header("❗ Missing Values")

    missing = df.isnull().sum()

    st.dataframe(missing[missing > 0])


def duplicate_analysis(df):
    st.header("📄 Duplicate Records")

    duplicates = df.duplicated().sum()

    st.metric("Duplicate Rows", duplicates)


def numeric_analysis(df):

    st.header("📈 Numeric Analysis")

    numeric = df.select_dtypes(include="number")

    if numeric.empty:
        st.info("No numeric columns found.")
        return

    st.subheader("Summary Statistics")
    st.dataframe(numeric.describe().T, use_container_width=True)

    st.divider()

    for col in numeric.columns:

        st.subheader(f"📊 {col}")

        c1, c2 = st.columns(2)

        with c1:

            fig = px.histogram(
                df,
                x=col,
                nbins=30,
                title=f"{col} Distribution"
            )

            fig.update_layout(
                template="plotly_white",
                height=350
            )

            st.plotly_chart(fig, use_container_width=True)

        with c2:

            fig = px.box(
                df,
                y=col,
                title=f"{col} Box Plot"
            )

            fig.update_layout(
                template="plotly_white",
                height=350
            )

            st.plotly_chart(fig, use_container_width=True)

        st.markdown("### 📌 Quick Insights")

        mean = df[col].mean()
        median = df[col].median()
        minimum = df[col].min()
        maximum = df[col].max()
        std = df[col].std()

        i1, i2, i3, i4, i5 = st.columns(5)

        i1.metric("Mean", round(mean,2))
        i2.metric("Median", round(median,2))
        i3.metric("Min", round(minimum,2))
        i4.metric("Max", round(maximum,2))
        i5.metric("Std", round(std,2))

        st.divider()

def category_analysis(df):

    st.header("📊 Category Analysis")

    cat_cols = df.select_dtypes(include=["object", "category"]).columns

    if len(cat_cols) == 0:
        st.info("No categorical columns found.")
        return

    for col in cat_cols:

        st.subheader(f"📌 {col}")

        value_counts = (
            df[col]
            .fillna("Missing")
            .value_counts()
            .reset_index()
        )

        value_counts.columns = [col, "Count"]

        c1, c2 = st.columns([1,1])

        with c1:

            st.markdown("### Top Categories")

            st.dataframe(
                value_counts.head(10),
                use_container_width=True
            )

        with c2:

            fig = px.bar(
                value_counts.head(10),
                x=col,
                y="Count",
                color="Count",
                title=f"{col} Distribution"
            )

            fig.update_layout(
                template="plotly_white",
                xaxis_title=col,
                yaxis_title="Count",
                height=400
            )

            st.plotly_chart(fig, use_container_width=True)

        if value_counts.shape[0] <= 8:

            fig = px.pie(
                value_counts,
                names=col,
                values="Count",
                hole=0.4,
                title=f"{col} Composition"
            )

            fig.update_layout(
                template="plotly_white",
                height=450
            )

            st.plotly_chart(fig, use_container_width=True)

        st.divider()

def correlation_analysis(df):

    st.header("🔥 Correlation Analysis")

    numeric = df.select_dtypes(include="number")

    if numeric.shape[1] < 2:
        st.info("At least two numeric columns are required.")
        return

    corr = numeric.corr()

    fig = px.imshow(
        corr,
        text_auto=".2f",
        color_continuous_scale="RdBu_r",
        aspect="auto",
        title="Correlation Heatmap"
    )

    fig.update_layout(
        template="plotly_white",
        height=650
    )

    st.plotly_chart(fig, use_container_width=True)

    st.subheader("Correlation Matrix")
    st.dataframe(corr.round(2), use_container_width=True)

    strong_corr = (
        corr.where(~pd.np.eye(corr.shape[0], dtype=bool))
            .stack()
            .reset_index()
    )

    strong_corr.columns = ["Column 1", "Column 2", "Correlation"]

    strong_corr["Absolute"] = strong_corr["Correlation"].abs()

    strong_corr = strong_corr.sort_values(
        "Absolute",
        ascending=False
    )

    st.subheader("Strongest Correlations")

    st.dataframe(
        strong_corr.head(10),
        use_container_width=True
        )

def outlier_analysis(df):

    st.header("🚨 Outlier Analysis")

    numeric = df.select_dtypes(include="number")

    if numeric.empty:
        st.info("No numeric columns found.")
        return

    outlier_data = []

    for col in numeric.columns:

        q1 = df[col].quantile(0.25)
        q3 = df[col].quantile(0.75)

        iqr = q3 - q1

        lower = q1 - (1.5 * iqr)
        upper = q3 + (1.5 * iqr)

        count = df[(df[col] < lower) | (df[col] > upper)].shape[0]

        outlier_data.append({
            "Column": col,
            "Outliers": count
        })

    outlier_df = pd.DataFrame(outlier_data)

    fig = px.bar(
        outlier_df,
        x="Column",
        y="Outliers",
        color="Outliers",
        text="Outliers",
        title="Outlier Count by Column"
    )

    fig.update_layout(
        template="plotly_white",
        height=500
    )

    st.plotly_chart(fig, use_container_width=True)

    st.dataframe(outlier_df, use_container_width=True)

    st.subheader("Box Plots")

    for col in numeric.columns:

        fig = px.box(
            df,
            y=col,
            points="outliers",
            title=f"{col} Outliers"
        )

        fig.update_layout(
            template="plotly_white",
            height=350
        )

        st.plotly_chart(fig, use_container_width=True)

def trend_analysis(df):

    st.header("📈 Trend Analysis")

    # Detect datetime columns
    date_cols = []

    for col in df.columns:
        try:
            converted = pd.to_datetime(df[col], errors="raise")
            date_cols.append(col)
        except:
            pass

    if len(date_cols) == 0:
        st.info("No Date Column Found.")
        return

    date_col = st.selectbox("Select Date Column", date_cols)

    numeric_cols = df.select_dtypes(include="number").columns

    if len(numeric_cols) == 0:
        st.warning("No Numeric Columns Available.")
        return

    value_col = st.selectbox("Select Value Column", numeric_cols)

    temp = df.copy()

    temp[date_col] = pd.to_datetime(temp[date_col])

    temp = temp.sort_values(date_col)

    freq = st.selectbox(
        "Group By",
        ["Daily", "Weekly", "Monthly", "Yearly"]
    )

    if freq == "Daily":
        trend = temp.groupby(temp[date_col].dt.date)[value_col].sum().reset_index()

    elif freq == "Weekly":
        trend = temp.groupby(temp[date_col].dt.to_period("W"))[value_col].sum().reset_index()
        trend[date_col] = trend[date_col].astype(str)

    elif freq == "Monthly":
        trend = temp.groupby(temp[date_col].dt.to_period("M"))[value_col].sum().reset_index()
        trend[date_col] = trend[date_col].astype(str)

    else:
        trend = temp.groupby(temp[date_col].dt.year)[value_col].sum().reset_index()

    fig = px.line(
        trend,
        x=date_col,
        y=value_col,
        markers=True,
        title=f"{value_col} Trend"
    )

    fig.update_layout(
        template="plotly_white",
        height=500,
        xaxis_title="Time",
        yaxis_title=value_col
    )

    st.plotly_chart(fig, use_container_width=True)

    st.subheader("Trend Data")
    st.dataframe(trend, use_container_width=True)
