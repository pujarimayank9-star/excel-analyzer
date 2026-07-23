import streamlit as st
import plotly.express as px

def show_bar_chart(df, column):

    fig = px.bar(
        df[column].value_counts().reset_index(),
        x="index",
        y=column,
        title=f"{column} Distribution"
    )

    st.plotly_chart(fig, use_container_width=True)


def show_pie_chart(df, column):

    fig = px.pie(
        df,
        names=column,
        title=f"{column} Share"
    )

    st.plotly_chart(fig, use_container_width=True)


def show_histogram(df, column):

    fig = px.histogram(
        df,
        x=column,
        title=f"{column} Histogram"
    )

    st.plotly_chart(fig, use_container_width=True)


def show_boxplot(df, column):

    fig = px.box(
        df,
        y=column,
        title=f"{column} Box Plot"
    )

    st.plotly_chart(fig, use_container_width=True)


def show_scatter(df, x, y):

    fig = px.scatter(
        df,
        x=x,
        y=y,
        title=f"{x} vs {y}"
    )

    st.plotly_chart(fig, use_container_width=True)


def show_line(df, column):

    fig = px.line(
        df,
        y=column,
        title=f"{column} Trend"
    )

    st.plotly_chart(fig, use_container_width=True)
