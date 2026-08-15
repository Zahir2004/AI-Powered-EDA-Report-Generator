"""
app.py
------
Streamlit entry point: AI-Powered EDA Report Generator.

Run with:
    streamlit run app.py
"""

import os
from datetime import datetime

import streamlit as st
from dotenv import load_dotenv

import eda_utils
import groq_utils

load_dotenv()  # loads GROQ_API_KEY from a local .env file into the environment

st.set_page_config(page_title="AI-Powered EDA Report Generator", layout="wide")

# The API key is configured in code only (via .env / environment variable) —
# there is no UI field for it, so it never has to be typed or pasted in the browser.
api_key = os.getenv("GROQ_API_KEY", "")
model_name = groq_utils.DEFAULT_MODEL  # fixed to the fastest, free-tier-friendly model

# --------------------------------------------------------------------------
# Sidebar: upload only
# --------------------------------------------------------------------------
st.sidebar.title("Upload")
uploaded_file = st.sidebar.file_uploader(
    "Upload a CSV or Excel file", type=["csv", "xlsx", "xls"]
)

st.title("📊 AI-Powered EDA Report Generator")
st.caption("Upload a CSV/Excel file to get an automated statistical EDA, "
           f"plus an AI-written narrative summary from Groq ({model_name}).")

if not api_key:
    st.error(
        "GROQ_API_KEY is not set. Copy `.env.example` to `.env` in the project "
        "folder and paste your key there (see README.md), then restart the app. "
        "The statistical EDA tabs below still work without a key — only the "
        "AI narrative section needs it."
    )

if not uploaded_file:
    st.info("Upload a file from the sidebar to get started.")
    st.stop()

# --------------------------------------------------------------------------
# Load data
# --------------------------------------------------------------------------
try:
    df = eda_utils.load_data(uploaded_file)
except Exception as e:
    st.error(f"Could not read file: {e}")
    st.stop()

if df.empty:
    st.warning("The uploaded file has no rows.")
    st.stop()

st.subheader("Preview")
st.dataframe(df.head(20), use_container_width=True)

# --------------------------------------------------------------------------
# Compute statistics (always local, no API calls, works offline)
# --------------------------------------------------------------------------
overview = eda_utils.get_overview(df)
missing_df = eda_utils.get_missing_summary(df)
numeric_summary_df = eda_utils.get_numeric_summary(df)
categorical_summary = eda_utils.get_categorical_summary(df)
corr_df = eda_utils.get_correlation_matrix(df)
outliers = eda_utils.detect_outliers_iqr(df)

tab_overview, tab_missing, tab_stats, tab_corr, tab_dist, tab_cat = st.tabs(
    ["Overview", "Missing Values", "Statistics", "Correlations", "Distributions", "Categorical"]
)

with tab_overview:
    c1, c2, c3, c4 = st.columns(4)
    c1.metric("Rows", overview["rows"])
    c2.metric("Columns", overview["columns"])
    c3.metric("Duplicate rows", overview["duplicate_rows"])
    c4.metric("Memory (MB)", overview["memory_mb"])
    st.write("**Numeric columns:**", ", ".join(overview["numeric_columns"]) or "None")
    st.write("**Categorical columns:**", ", ".join(overview["categorical_columns"]) or "None")

with tab_missing:
    if missing_df.empty:
        st.success("No missing values detected.")
    else:
        st.dataframe(missing_df, use_container_width=True)
        st.bar_chart(missing_df["missing_pct"])

with tab_stats:
    if numeric_summary_df.empty:
        st.info("No numeric columns to summarize.")
    else:
        st.dataframe(numeric_summary_df, use_container_width=True)
    if outliers:
        st.write("**Potential outliers (IQR rule):**")
        st.json(outliers)
    else:
        st.success("No significant outliers detected via the IQR rule.")

with tab_corr:
    if corr_df.empty:
        st.info("Need at least two numeric columns for a correlation matrix.")
        corr_fig = None
    else:
        st.dataframe(corr_df, use_container_width=True)
        corr_fig = eda_utils.generate_correlation_heatmap(corr_df)
        st.pyplot(corr_fig)

with tab_dist:
    hist_figs = eda_utils.generate_numeric_histograms(df)
    if not hist_figs:
        st.info("No numeric columns to plot.")
    else:
        cols = st.columns(2)
        for i, (col_name, fig) in enumerate(hist_figs.items()):
            with cols[i % 2]:
                st.pyplot(fig)

with tab_cat:
    bar_figs = eda_utils.generate_categorical_bars(df)
    if not bar_figs:
        st.info("No categorical columns to plot.")
    else:
        cols = st.columns(2)
        for i, (col_name, fig) in enumerate(bar_figs.items()):
            with cols[i % 2]:
                st.pyplot(fig)

# --------------------------------------------------------------------------
# AI narrative (only computed stats are sent to Groq, never raw rows)
# --------------------------------------------------------------------------
st.markdown("---")
st.subheader("🤖 AI Narrative Insights")

if "narrative" not in st.session_state:
    st.session_state.narrative = None

generate_clicked = st.button("Generate AI Report", type="primary")

if generate_clicked:
    if not api_key:
        st.warning("GROQ_API_KEY is not set — see the message above.")
    else:
        with st.spinner(f"Asking Groq ({model_name}) to write the report..."):
            try:
                client = groq_utils.get_client(api_key)
                summary_text = groq_utils.build_eda_summary_text(
                    overview, missing_df, numeric_summary_df,
                    categorical_summary, corr_df, outliers,
                )
                st.session_state.narrative = groq_utils.generate_narrative_report(
                    client, summary_text, model=model_name
                )
            except Exception as e:
                st.error(f"Groq request failed: {e}")
                st.session_state.narrative = None

if st.session_state.narrative:
    st.markdown(st.session_state.narrative)

    st.markdown("---")
    st.subheader("⬇️ Download Report")

    generated_at = datetime.now().strftime("%Y-%m-%d %H:%M")
    md_report = (
        f"# EDA Report — {uploaded_file.name}\n"
        f"_Generated {generated_at} using {model_name}_\n\n"
        f"{st.session_state.narrative}\n\n"
        "---\n## Computed Statistics\n\n"
        f"```\n{groq_utils.build_eda_summary_text(overview, missing_df, numeric_summary_df, categorical_summary, corr_df, outliers)}\n```\n"
    )

    st.download_button(
        "Download Markdown report",
        data=md_report,
        file_name=f"eda_report_{uploaded_file.name.rsplit('.', 1)[0]}.md",
        mime="text/markdown",
    )
