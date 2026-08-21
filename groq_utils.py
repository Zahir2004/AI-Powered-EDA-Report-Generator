"""
groq_utils.py
-------------
Wraps the Groq API (groq SDK, OpenAI-compatible chat completions) to turn
computed EDA statistics into a plain-English narrative report.

Groq was chosen after repeated account-level access problems on the Gemini
API (invalid "AQ."-format keys, a retired model, and a project-level 403
"denied access" - all confirmed as Google-side issues, not code bugs).
Groq's free tier has no known equivalent problems as of writing and is
also one of the fastest inference providers available.

Design choice: we NEVER send the raw dataset to Groq. Only the
already-computed summary statistics (shapes, describe() tables, missing
%, correlations, outlier counts) are sent. This keeps token usage small,
keeps cost bounded regardless of file size, and avoids sending
potentially sensitive row-level data to a third-party API.
"""

from groq import Groq

# llama-3.3-70b-versatile: Groq's well-established, fast, free-tier model -
# a strong default for narrative/report writing. If you want more speed
# and don't mind slightly less nuance, swap to "llama-3.1-8b-instant".
DEFAULT_MODEL = "openai/gpt-oss-120b"

SYSTEM_INSTRUCTION = (
    "You are a senior data analyst writing the narrative section of an "
    "exploratory data analysis (EDA) report. You are given ONLY pre-computed "
    "summary statistics for a tabular dataset -- never the raw rows. "
    "Write a clear, well-structured Markdown report with these sections: "
    "## Dataset Overview, ## Key Observations, ## Data Quality Issues, "
    "## Notable Relationships, and ## Recommended Next Steps. "
    "Be specific and reference actual column names and numbers from the "
    "statistics provided. Do not invent values that are not present in the "
    "input. Keep it concise: prefer bullet points over long paragraphs."
)


def get_client(api_key: str) -> Groq:
    if not api_key:
        raise ValueError("A Groq API key is required.")
    return Groq(api_key=api_key)


def build_eda_summary_text(
    overview: dict,
    missing_df,
    numeric_summary_df,
    categorical_summary: dict,
    corr_df,
    outliers: dict,
) -> str:
    """Serialize computed statistics into a compact plain-text block."""
    lines = []

    lines.append("### Dataset shape")
    lines.append(f"- Rows: {overview['rows']}")
    lines.append(f"- Columns: {overview['columns']}")
    lines.append(f"- Duplicate rows: {overview['duplicate_rows']}")
    lines.append(f"- Numeric columns: {', '.join(overview['numeric_columns']) or 'none'}")
    lines.append(f"- Categorical columns: {', '.join(overview['categorical_columns']) or 'none'}")

    lines.append("\n### Missing values (column: count, pct)")
    if missing_df is not None and not missing_df.empty:
        for col, row in missing_df.iterrows():
            lines.append(f"- {col}: {int(row['missing_count'])} ({row['missing_pct']}%)")
    else:
        lines.append("- No missing values detected.")

    lines.append("\n### Numeric column statistics (describe())")
    if numeric_summary_df is not None and not numeric_summary_df.empty:
        lines.append(numeric_summary_df.to_string())
    else:
        lines.append("- No numeric columns.")

    lines.append("\n### Categorical column summaries (top values)")
    if categorical_summary:
        for col, info in categorical_summary.items():
            top_vals = ", ".join(f"{k}={v}" for k, v in info["top_values"].items())
            lines.append(f"- {col}: {info['unique_values']} unique values; top: {top_vals}")
    else:
        lines.append("- No categorical columns.")

    lines.append("\n### Correlation matrix (numeric columns)")
    if corr_df is not None and not corr_df.empty:
        lines.append(corr_df.to_string())
    else:
        lines.append("- Not enough numeric columns for correlation.")

    lines.append("\n### Outliers (IQR rule, column: count / pct of non-null values)")
    if outliers:
        for col, info in outliers.items():
            lines.append(f"- {col}: {info['count']} ({info['pct']}%)")
    else:
        lines.append("- No significant outliers detected.")

    return "\n".join(lines)


def generate_narrative_report(client: Groq, eda_summary_text: str, model: str = DEFAULT_MODEL) -> str:
    """Call Groq with the summary stats and return a Markdown narrative."""
    completion = client.chat.completions.create(
        model=model,
        messages=[
            {"role": "system", "content": SYSTEM_INSTRUCTION},
            {
                "role": "user",
                "content": (
                    "Here are the computed EDA statistics for a dataset. "
                    "Write the narrative report described in your instructions.\n\n"
                    f"{eda_summary_text}"
                ),
            },
        ],
        temperature=0.3,
    )
    return completion.choices[0].message.content
