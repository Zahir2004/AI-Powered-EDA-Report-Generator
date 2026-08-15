"""
eda_utils.py
------------
Pure-pandas/matplotlib EDA helpers. Nothing here calls any external API,
so all of this works even without a Gemini key.
"""

import base64
import io

import matplotlib
matplotlib.use("Agg")  # headless backend, required for Streamlit
import matplotlib.pyplot as plt
import numpy as np
import pandas as pd
import seaborn as sns

sns.set_theme(style="whitegrid")

MAX_NUMERIC_PLOTS = 12       # cap how many histograms we generate
MAX_CATEGORICAL_PLOTS = 8    # cap how many bar charts we generate
TOP_N_CATEGORIES = 15        # bar chart shows at most this many categories


# --------------------------------------------------------------------------
# Loading
# --------------------------------------------------------------------------
def load_data(uploaded_file) -> pd.DataFrame:
    """Load a Streamlit UploadedFile (csv/xlsx/xls) into a DataFrame."""
    name = uploaded_file.name.lower()
    if name.endswith(".csv"):
        return pd.read_csv(uploaded_file)
    if name.endswith((".xlsx", ".xls")):
        return pd.read_excel(uploaded_file)
    raise ValueError(f"Unsupported file type: {uploaded_file.name}")


# --------------------------------------------------------------------------
# Summaries
# --------------------------------------------------------------------------
def get_overview(df: pd.DataFrame) -> dict:
    return {
        "rows": df.shape[0],
        "columns": df.shape[1],
        "duplicate_rows": int(df.duplicated().sum()),
        "memory_mb": round(df.memory_usage(deep=True).sum() / (1024 ** 2), 3),
        "numeric_columns": df.select_dtypes(include=np.number).columns.tolist(),
        "categorical_columns": df.select_dtypes(exclude=np.number).columns.tolist(),
    }


def get_missing_summary(df: pd.DataFrame) -> pd.DataFrame:
    missing = df.isna().sum()
    pct = (missing / len(df) * 100).round(2)
    out = pd.DataFrame({"missing_count": missing, "missing_pct": pct})
    return out[out["missing_count"] > 0].sort_values("missing_count", ascending=False)


def get_numeric_summary(df: pd.DataFrame) -> pd.DataFrame:
    numeric_df = df.select_dtypes(include=np.number)
    if numeric_df.empty:
        return pd.DataFrame()
    return numeric_df.describe().T.round(3)


def get_categorical_summary(df: pd.DataFrame, top_n: int = 5) -> dict:
    cat_df = df.select_dtypes(exclude=np.number)
    summary = {}
    for col in cat_df.columns:
        vc = cat_df[col].value_counts(dropna=False).head(top_n)
        summary[col] = {
            "unique_values": int(cat_df[col].nunique(dropna=True)),
            "top_values": vc.to_dict(),
        }
    return summary


def get_correlation_matrix(df: pd.DataFrame) -> pd.DataFrame:
    numeric_df = df.select_dtypes(include=np.number)
    if numeric_df.shape[1] < 2:
        return pd.DataFrame()
    return numeric_df.corr().round(3)


def detect_outliers_iqr(df: pd.DataFrame) -> dict:
    """Simple IQR-rule outlier count per numeric column."""
    numeric_df = df.select_dtypes(include=np.number)
    outliers = {}
    for col in numeric_df.columns:
        series = numeric_df[col].dropna()
        if series.empty:
            continue
        q1, q3 = series.quantile(0.25), series.quantile(0.75)
        iqr = q3 - q1
        if iqr == 0:
            continue
        lower, upper = q1 - 1.5 * iqr, q3 + 1.5 * iqr
        count = int(((series < lower) | (series > upper)).sum())
        if count > 0:
            outliers[col] = {
                "count": count,
                "pct": round(count / len(series) * 100, 2),
            }
    return outliers


# --------------------------------------------------------------------------
# Plots (matplotlib figures, ready for st.pyplot or HTML embedding)
# --------------------------------------------------------------------------
def generate_numeric_histograms(df: pd.DataFrame) -> dict:
    numeric_cols = df.select_dtypes(include=np.number).columns.tolist()[:MAX_NUMERIC_PLOTS]
    figs = {}
    for col in numeric_cols:
        fig, ax = plt.subplots(figsize=(5, 3.2))
        sns.histplot(df[col].dropna(), kde=True, ax=ax, color="#4C78A8")
        ax.set_title(f"Distribution: {col}")
        fig.tight_layout()
        figs[col] = fig
    return figs


def generate_categorical_bars(df: pd.DataFrame) -> dict:
    cat_cols = df.select_dtypes(exclude=np.number).columns.tolist()[:MAX_CATEGORICAL_PLOTS]
    figs = {}
    for col in cat_cols:
        vc = df[col].value_counts(dropna=False).head(TOP_N_CATEGORIES)
        fig, ax = plt.subplots(figsize=(5, 3.2))
        sns.barplot(x=vc.values, y=vc.index.astype(str), ax=ax, color="#F58518")
        ax.set_title(f"Top values: {col}")
        ax.set_xlabel("Count")
        fig.tight_layout()
        figs[col] = fig
    return figs


def generate_correlation_heatmap(corr: pd.DataFrame):
    if corr.empty:
        return None
    fig, ax = plt.subplots(figsize=(min(10, 1 + len(corr.columns) * 0.8), min(8, 1 + len(corr.columns) * 0.7)))
    sns.heatmap(corr, annot=len(corr.columns) <= 15, cmap="coolwarm", center=0, ax=ax, fmt=".2f")
    ax.set_title("Correlation Heatmap")
    fig.tight_layout()
    return fig


def fig_to_base64(fig) -> str:
    """Encode a matplotlib figure as a base64 PNG string for HTML embedding."""
    buf = io.BytesIO()
    fig.savefig(buf, format="png", dpi=110, bbox_inches="tight")
    buf.seek(0)
    return base64.b64encode(buf.read()).decode("utf-8")
