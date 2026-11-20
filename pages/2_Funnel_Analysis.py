"""
pages/2_Funnel_Analysis.py - Multi-Stage Funnel & Conversion Leakage (UI-03)
"""

import streamlit as st
import pandas as pd
import numpy as np
import plotly.express as px
import plotly.graph_objects as go

from src.ui import render_header, render_kpi_cards, render_insight_box
from src.data_loader import load_campaigns_data, get_channel_list

st.set_page_config(page_title="Funnel Analysis", page_icon="🔻", layout="wide")

def main():
    render_header(
        title="Multi-Stage Conversion Funnel & Leakage Analyzer",
        subtitle="Track prospective customer drop-off from impression to checkout across marketing channels",
        badge_text="UI-03 Analytics"
    )

    df_camp = load_campaigns_data()
    channel_list = get_channel_list()

    # Channel Selector
    st.sidebar.title("🎛️ Funnel Controls")
    selected_channel = st.sidebar.selectbox("Focal Acquisition Channel", options=channel_list, index=1) # Default Social

    ch_df = df_camp[df_camp["channel_name"] == selected_channel]
    
    total_spend = ch_df["spend_amount"].sum()
    total_impr = ch_df["impressions"].sum()
    total_clicks = ch_df["clicks"].sum()
    total_leads = ch_df["leads"].sum()
    total_conv = ch_df["conversions_count"].sum()

    cpc = round(total_spend / max(total_clicks, 1), 2)
    cpl = round(total_spend / max(total_leads, 1), 2)
    cac = round(total_spend / max(total_conv, 1), 2)

    # Top Metrics
    render_kpi_cards([
        {"title": f"{selected_channel} Spend", "value": f"${total_spend:,.0f}", "sub": "Total Outlay", "status": "neutral"},
        {"title": "Cost Per Click (CPC)", "value": f"${cpc:.2f}", "sub": "Top Funnel", "status": "positive" if cpc <= 2.0 else "neutral"},
        {"title": "Cost Per Lead (CPL)", "value": f"${cpl:.2f}", "sub": "Mid Funnel", "status": "neutral"},
        {"title": "Final CAC", "value": f"${cac:.2f}", "sub": "Bottom Funnel", "status": "positive" if cac <= 50.0 else "critical"}
    ])

    st.markdown("<div style='height: 12px;'></div>", unsafe_allow_html=True)

    # Funnel Chart
    st.subheader(f"🔻 Multi-Stage Funnel Attrition: {selected_channel}")
    
    funnel_stages = ["1. Ad Impressions", "2. Click-Throughs", "3. Qualified Leads", "4. Paying Customers"]
    stage_values = [total_impr, total_clicks, total_leads, total_conv]

    fig_funnel = go.Figure(go.Funnel(
        y=funnel_stages,
        x=stage_values,
        textinfo="value+percent previous+percent initial",
        marker=dict(color=["#1E3A8A", "#2563EB", "#60A5FA", "#10B981"])
    ))

    fig_funnel.update_layout(
        margin=dict(l=20, r=20, t=20, b=20),
        height=380,
        plot_bgcolor="#FFFFFF",
        paper_bgcolor="#FFFFFF"
    )
    st.plotly_chart(fig_funnel, use_container_width=True)

    # Conversion Velocity Diagnosis
    click_to_lead_rate = round((total_leads / max(total_clicks, 1)) * 100.0, 2)
    lead_to_sale_rate = round((total_conv / max(total_leads, 1)) * 100.0, 2)

    c1, c2 = st.columns(2)
    with c1:
        render_insight_box(
            "info",
            "Mid-Funnel Capture Velocity",
            f"{selected_channel} converts {click_to_lead_rate}% of ad clicks into qualified leads. Benchmark for B2B SaaS is 10–15%."
        )
    with c2:
        if lead_to_sale_rate < 5.0:
            render_insight_box(
                "critical",
                "Bottom-Funnel Leakage Alert",
                f"Severe leakage detected at final conversion stage: only {lead_to_sale_rate}% of leads convert to paying customers. Prospective clients drop off during the checkout or pricing gate."
            )
        else:
            render_insight_box(
                "success",
                "Healthy Conversion Velocity",
                f"Strong bottom-of-funnel velocity: {lead_to_sale_rate}% of captured leads convert into active paying accounts."
            )

if __name__ == "__main__":
    main()
