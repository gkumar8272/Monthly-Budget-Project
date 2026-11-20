"""
pages/1_Channel_CAC.py - Channel CAC vs Blended Benchmark & Payback Trends (UI-02)
"""

import streamlit as st
import pandas as pd
import numpy as np
import plotly.express as px
import plotly.graph_objects as go

from src.ui import render_header, render_kpi_cards, render_insight_box
from src.data_loader import load_campaigns_data, load_customers_data, get_channel_list
from src.cac_calculator import calculate_channel_cac_summary

st.set_page_config(page_title="Channel CAC & Payback", page_icon="📊", layout="wide")

def main():
    render_header(
        title="Channel-Specific Customer Acquisition Cost (CAC)",
        subtitle="Deconstruct blended averages into discrete channel unit economics and payback benchmarks",
        badge_text="UI-02 Analytics"
    )

    df_camp = load_campaigns_data()
    df_cust = load_customers_data()
    channel_list = get_channel_list()

    # Sidebar Filter
    st.sidebar.title("🎛️ Channel Filters")
    selected_channels = st.sidebar.multiselect("Active Channels", options=channel_list, default=channel_list)

    filtered_camp = df_camp[df_camp["channel_name"].isin(selected_channels)]
    if filtered_camp.empty:
        st.warning("Please select at least one channel.")
        return

    # Compute Summary
    summary = calculate_channel_cac_summary(filtered_camp)
    blended_cac = round(filtered_camp["spend_amount"].sum() / max(filtered_camp["conversions_count"].sum(), 1), 2)

    # 1. Bar Chart: Channel CAC vs Blended Line
    st.subheader("📊 Channel-Specific CAC vs Blended Average Reference")
    fig_bar = go.Figure()

    fig_bar.add_trace(go.Bar(
        x=summary["channel_name"],
        y=summary["channel_cac"],
        name="Channel CAC ($)",
        marker_color=["#2563EB", "#F59E0B", "#10B981", "#6366F1", "#8B5CF6"][:len(summary)],
        text=summary["channel_cac"].apply(lambda v: f"${v:,.2f}"),
        textposition="outside"
    ))

    fig_bar.add_hline(
        y=blended_cac,
        line_dash="dash",
        line_color="#DC2626",
        annotation_text=f"Blended Average CAC (${blended_cac:,.2f})",
        annotation_position="top right"
    )

    fig_bar.update_layout(
        xaxis_title="Acquisition Channel",
        yaxis_title="Cost Per Customer ($)",
        plot_bgcolor="#FFFFFF",
        paper_bgcolor="#FFFFFF",
        height=380
    )
    st.plotly_chart(fig_bar, use_container_width=True)

    # 2. Payback Horizon Table & Alert
    st.subheader("⏱️ Capital Payback Periods per Channel")
    filtered_cust = df_cust[df_cust["channel_name"].isin(selected_channels)]
    
    payback_table = filtered_cust.groupby("channel_name").agg(
        total_customers=("customer_id", "count"),
        avg_mrr=("monthly_revenue", "mean"),
        avg_cac=("cac_attributed", "mean"),
        avg_payback=("payback_months", "mean")
    ).reset_index()

    payback_table["avg_payback"] = payback_table["avg_payback"].round(1)
    payback_table["avg_cac"] = payback_table["avg_cac"].round(2)
    payback_table["avg_mrr"] = payback_table["avg_mrr"].round(2)
    payback_table["solvency_status"] = np.where(
        payback_table["avg_payback"] <= 12.0,
        "✅ Safe (<12 Mo)",
        "🚨 Critical (>12 Mo)"
    )

    col1, col2 = st.columns([7, 5])
    with col1:
        st.dataframe(
            payback_table.rename(columns={
                "channel_name": "Channel",
                "total_customers": "Customers",
                "avg_mrr": "Avg MRR ($)",
                "avg_cac": "Avg CAC ($)",
                "avg_payback": "Payback (Months)",
                "solvency_status": "Solvency Status"
            }),
            use_container_width=True,
            hide_index=True
        )

    with col2:
        social_payback = payback_table[payback_table["channel_name"] == "Social"]
        if not social_payback.empty and social_payback["avg_payback"].iloc[0] > 12.0:
            render_insight_box(
                "critical",
                "Payback Threshold Breach Alert",
                f"Paid Social is running at an average payback period of {social_payback['avg_payback'].iloc[0]} months, exceeding the enterprise 12-month solvency ceiling. Review ad creatives and target audience qualification."
            )
        else:
            render_insight_box(
                "success",
                "Working Capital Stability",
                "Selected channels recoup customer acquisition spend in under 12 months, preserving enterprise working capital runway."
            )

if __name__ == "__main__":
    main()
