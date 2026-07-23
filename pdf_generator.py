from reportlab.platypus import SimpleDocTemplate, Paragraph
from reportlab.lib.styles import getSampleStyleSheet
import streamlit as st
import os

def pdf_download(df):

    st.header("📑 PDF Report")

    if st.button("Generate PDF"):

        filename = "AI_Data_Report.pdf"

        doc = SimpleDocTemplate(filename)

        styles = getSampleStyleSheet()

        story = []

        story.append(
            Paragraph("<b>AI Smart Excel Analysis Report</b>", styles["Title"])
        )

        story.append(
            Paragraph(f"Rows : {df.shape[0]}", styles["Normal"])
        )

        story.append(
            Paragraph(f"Columns : {df.shape[1]}", styles["Normal"])
        )

        story.append(
            Paragraph(
                f"Missing Values : {df.isnull().sum().sum()}",
                styles["Normal"]
            )
        )

        story.append(
            Paragraph(
                f"Duplicate Rows : {df.duplicated().sum()}",
                styles["Normal"]
            )
        )

        story.append(
            Paragraph("<br/><b>Column Names</b>", styles["Heading2"])
        )

        for col in df.columns:
            story.append(
                Paragraph(str(col), styles["Normal"])
            )

        doc.build(story)

        with open(filename, "rb") as f:

            st.download_button(
                "⬇ Download PDF",
                f,
                file_name=filename
            )

        os.remove(filename)
