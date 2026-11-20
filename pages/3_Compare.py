"""
pages/3_Compare.py - Volume vs Cost-Efficiency 2x2 Quadrant Matrix (UI-04)
"""

import streamlit as st
import pandas as pd
import numpy as np
import plotly.express as px

from src.ui import render_header, render_insight_box
from src.data_loader import load_campaigns_data
from src.cac_calculator import calculate_channel_cac_summary

st.set_page_config(page_title="Channel Quadrant Compare", page_icon="⚖️", layout="wide")

def main():
    render_header(
        title="Acquisition Volume vs. Cost-Efficiency Quadrant Matrix",
        subtitle="Evaluate channel trade-offs across acquisition volume and unit acquisition cost",
        badge_text="UI-04 Strategic Compare"
    )

    df_camp = load_campaigns_data()
    summary = calculate_channel_cac_summary(df_camp)

    # 2x2 Scatter Matrix
    median_vol = summary["total_conversions"].median()
    median_cac = summary["channel_cac"].median()

    st.subheader("🎯 Channel Strategic Quadrant Matrix")
    
    fig = px.scatter(
        summary,
        x="total_conversions",
        y="channel_cac",
        size="total_spend",
        color="channel_name",
        text="channel_name",
        hover_data={"total_spend": ":$,.0f", "total_conversions": ":,", "channel_cac": ":$,.2f"},
        size_max=35,
        color_discrete_map={
            "Search": "#2563EB",
            "Social": "#F59E0B",
            "Referral": "#10B981",
            "Email": "#6366F1",
            "Organic": "#8B5CF6"
        }
    )

    # Add Quadrant Reference Lines
    fig.add_vline(x=median_vol, line_dash="dash", line_color="#94A3B8")
    fig.add_hline(y=median_cac, line_dash="dash", line_color="#94A3B8")

    # Add Quadrant Annotations
    max_x = summary["total_conversions"].max() * 1.05
    max_y = summary["channel_cac"].max() * 1.05
    min_x = summary["total_conversions"].min() * 0.85
    min_y = summary["channel_cac"].min() * 0.85

    fig.add_annotation(x=max_x, y=min_y, text="🌟 SCALE WINNERS<br>(High Volume, Low CAC)", showarrow=False, font=dict(color="#166534", size=10))
    fig.add_annotation(x=min_x, y=min_y, text="💎 NICHE CHAMPIONS<br>(Low Volume, Low CAC)", showarrow=False, font=dict(color="#1E40AF", size=10))
    fig.add_annotation(x=max_x, y=max_y, text="⚠️ EXPENSIVE SCALERS<br>(High Volume, High CAC)", showarrow=False, font=dict(color="#9A3412", size=10))
    fig.add_annotation(x=min_x, y=max_y, text="🛑 UNDERPERFORMERS<br>(Low Volume, High CAC)", showarrow=False, font=dict(color="#991B1B", size=10))

    fig.update_traces(textposition="top center")
    fig.update_layout(
        xaxis_title="Total Acquired Customers (Volume)",
        yaxis_title="Customer Acquisition Cost - CAC ($)",
        plot_bgcolor="#FFFFFF",
        paper_bgcolor="#FFFFFF",
        height=450
    )
    st.plotly_chart(fig, use_container_width=True)

    # Strategic Matrix Table
    st.subheader("📋 Quadrant Classification Summary")
    
    def classify_quadrant(row):
        vol = row["total_conversions"]
        cac = row["channel_cac"]
        if vol >= median_vol and cac <= median_cac:
            return "🌟 Scale Winner"
        elif vol < median_vol and cac <= median_cac:
            return "💎 Niche Champion"
        elif vol >= median_vol and cac > median_cac:
            return "⚠️ Expensive Scaler"
        else:
            return "🛑 Underperformer"

    summary["strategic_quadrant"] = summary.apply(classify_quadrant, axis=1)

    st.dataframe(
        summary[["channel_name", "total_conversions", "total_spend", "channel_cac", "strategic_quadrant"]].rename(columns={
            "channel_name": "Channel",
            "total_conversions": "Total Customers",
            "total_spend": "Total Spend ($)",
            "channel_cac": "Channel CAC ($)",
            "strategic_quadrant": "Strategic Classification"
        }),
        use_container_width=True,
        hide_index=True
    )

if __name__ == "__main__":
    main()
