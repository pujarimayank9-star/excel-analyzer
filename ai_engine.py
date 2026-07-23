import streamlit as st
from google import genai


def ai_business_analysis(df):

    st.header("🤖 AI Business Analyst")

    try:

        client = genai.Client(
            api_key=st.secrets["GEMINI_API_KEY"]
        )

    except Exception:

        st.error(
            "Gemini API Key not found.\n\n"
            "Add GEMINI_API_KEY inside Streamlit Secrets."
        )

        return

    if st.button("🚀 Analyze with AI"):

        with st.spinner("Gemini is analysing your dataset..."):

            prompt = f"""
You are an expert Data Analyst and Business Consultant.

Analyse this dataset carefully.

Dataset Shape:
Rows = {df.shape[0]}
Columns = {df.shape[1]}

Columns:
{list(df.columns)}

Statistics:

{df.describe(include='all').to_string()}

Missing Values:

{df.isnull().sum().to_string()}

Now provide a professional report.

Use the following format.

# Executive Summary

# Important Insights

# Risks

# Opportunities

# Business Recommendations

# Conclusion

Explain everything in simple English.
"""

            try:

                response = client.models.generate_content(
                    model="gemini-2.5-flash",
                    contents=prompt
                )

                st.success("Analysis Completed")

                st.markdown(response.text)

            except Exception as e:

                st.error(e)
