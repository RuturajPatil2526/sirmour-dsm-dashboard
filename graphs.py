"""
graphs.py
=========
All Plotly chart builders for the AI Schedule Evaluation Dashboard.

Every function takes the fully-computed per-block DataFrame (the output of
calculator.evaluate_schedule) and returns a go.Figure, rendered live in the
Streamlit "Charts" tab.
"""

from __future__ import annotations

import plotly.graph_objects as go
import pandas as pd

# A small, consistent, colour-blind-friendly palette tied to the dashboard's
# green & white theme. Semantic colours (over/under generation) are kept
# distinct from the brand green used for the primary "Predicted" series.
DARK_GREEN = "#0F3D2E"
COLOR_PREDICTED = "#15803D"   # brand green (AI Schedule / Predicted)
COLOR_ACTUAL = "#F59E0B"      # amber (Actual -- contrasts against green)
COLOR_OVER = "#16A34A"        # green (over-generation)
COLOR_UNDER = "#DC2626"       # red (under-generation)
COLOR_PENALTY = "#7C3AED"     # violet (kept distinct so penalty stands out)
COLOR_CUMULATIVE = "#0F766E"  # teal-green
COLOR_NEUTRAL = "#64748B"     # slate

TEMPLATE = "plotly_white"


def _apply_theme(fig: go.Figure) -> go.Figure:
    """Apply the dashboard's green & white visual theme consistently across
    every chart: white canvas, mint gridlines, dark-green titles."""
    fig.update_layout(
        paper_bgcolor="#FFFFFF",
        plot_bgcolor="#FFFFFF",
        font=dict(color="#0F172A", family="Segoe UI, Inter, sans-serif"),
        title_font=dict(color=DARK_GREEN, size=17, family="Segoe UI, Inter, sans-serif"),
        legend=dict(bgcolor="rgba(0,0,0,0)"),
        margin=dict(t=60, l=10, r=10, b=10),
    )
    fig.update_xaxes(gridcolor="#E6F4EA", zerolinecolor="#BBF7D0", linecolor="#BBF7D0")
    fig.update_yaxes(gridcolor="#E6F4EA", zerolinecolor="#BBF7D0", linecolor="#BBF7D0")
    return fig


def _x_labels(df: pd.DataFrame):
    return df["Block"].astype(str) + " (" + df["Time_Label"].astype(str) + ")"


def fig_actual_vs_predicted(df: pd.DataFrame) -> go.Figure:
    fig = go.Figure()
    fig.add_trace(go.Scatter(
        x=df["Block"], y=df["Scheduled_MW"], mode="lines+markers", name="Predicted (Scheduled) MW",
        line=dict(color=COLOR_PREDICTED, width=2), marker=dict(size=4),
        hovertext=df["Time_Label"], hovertemplate="Block %{x} (%{hovertext})<br>Scheduled: %{y:.3f} MW<extra></extra>",
    ))
    fig.add_trace(go.Scatter(
        x=df["Block"], y=df["Actual_MW"], mode="lines+markers", name="Actual MW",
        line=dict(color=COLOR_ACTUAL, width=2), marker=dict(size=4),
        hovertext=df["Time_Label"], hovertemplate="Block %{x} (%{hovertext})<br>Actual: %{y:.3f} MW<extra></extra>",
    ))
    fig.update_layout(
        title="Actual vs Predicted Generation (per 15-min Block)",
        xaxis_title="Block Number", yaxis_title="MW", template=TEMPLATE,
        legend=dict(orientation="h", yanchor="bottom", y=1.02, xanchor="right", x=1),
        hovermode="x unified",
    )
    return fig


def fig_deviation_bar(df: pd.DataFrame) -> go.Figure:
    colors = [COLOR_OVER if v >= 0 else COLOR_UNDER for v in df["Deviation_MW"]]
    fig = go.Figure(go.Bar(
        x=df["Block"], y=df["Deviation_MW"], marker_color=colors,
        hovertext=df["Time_Label"],
        hovertemplate="Block %{x} (%{hovertext})<br>Deviation: %{y:.3f} MW<extra></extra>",
    ))
    fig.add_hline(y=0, line_color=COLOR_NEUTRAL, line_width=1)
    fig.update_layout(
        title="Deviation (MW) by Block", xaxis_title="Block Number",
        yaxis_title="Deviation MW", template=TEMPLATE,
    )
    return fig


def fig_penalty_bar(df: pd.DataFrame) -> go.Figure:
    fig = go.Figure(go.Bar(
        x=df["Block"], y=df["Total_Block_Penalty"], marker_color=COLOR_PENALTY,
        hovertext=df["Time_Label"],
        hovertemplate="Block %{x} (%{hovertext})<br>Penalty: Rs %{y:,.2f}<extra></extra>",
    ))
    fig.update_layout(
        title="DSM Penalty (Rs) by Block", xaxis_title="Block Number",
        yaxis_title="Penalty (Rs)", template=TEMPLATE,
    )
    return fig


def fig_cumulative_penalty(df: pd.DataFrame) -> go.Figure:
    cum = df["Total_Block_Penalty"].cumsum()
    fig = go.Figure(go.Scatter(
        x=df["Block"], y=cum, mode="lines+markers", fill="tozeroy",
        line=dict(color=COLOR_CUMULATIVE, width=2), marker=dict(size=4),
        hovertext=df["Time_Label"],
        hovertemplate="Block %{x} (%{hovertext})<br>Cumulative Penalty: Rs %{y:,.2f}<extra></extra>",
    ))
    fig.update_layout(
        title="Cumulative DSM Penalty Through the Day", xaxis_title="Block Number",
        yaxis_title="Cumulative Penalty (Rs)", template=TEMPLATE,
    )
    return fig


def fig_over_under_pie(df: pd.DataFrame) -> go.Figure:
    counts = df["Deviation_Type"].value_counts()
    labels = list(counts.index)
    values = list(counts.values)
    color_map = {"Over Generation": COLOR_OVER, "Under Generation": COLOR_UNDER, "No Deviation": COLOR_NEUTRAL}
    colors = [color_map.get(l, COLOR_NEUTRAL) for l in labels]
    fig = go.Figure(go.Pie(
        labels=labels, values=values, marker=dict(colors=colors), hole=0.45,
        textinfo="label+percent",
    ))
    fig.update_layout(title="Over Generation vs Under Generation Blocks", template=TEMPLATE)
    return fig


def fig_deviation_histogram(df: pd.DataFrame) -> go.Figure:
    fig = go.Figure(go.Histogram(
        x=df["Deviation_%"], marker_color=COLOR_PREDICTED, nbinsx=20,
    ))
    fig.update_layout(
        title="Distribution of Deviation %", xaxis_title="Deviation % of Capacity",
        yaxis_title="Number of Blocks", template=TEMPLATE, bargap=0.05,
    )
    return fig


def all_figures(df: pd.DataFrame) -> "dict[str, go.Figure]":
    """Convenience helper returning every chart (themed) keyed by a short
    slug -- used by both app.py and excel_export.py so the two never drift
    apart."""
    return {
        "actual_vs_predicted": _apply_theme(fig_actual_vs_predicted(df)),
        "deviation_bar": _apply_theme(fig_deviation_bar(df)),
        "penalty_bar": _apply_theme(fig_penalty_bar(df)),
        "cumulative_penalty": _apply_theme(fig_cumulative_penalty(df)),
        "over_under_pie": _apply_theme(fig_over_under_pie(df)),
        "deviation_histogram": _apply_theme(fig_deviation_histogram(df)),
    }
