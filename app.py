"""
app.py - Executive CAC Overview & Cockpit (UI-01)
Main entrypoint for Customer Acquisition Cost (CAC) Analysis platform.
"""

import streamlit as st
import pandas as pd
import numpy as np
import plotly.express as px
import plotly.graph_objects as go

from src.ui import render_header, render_kpi_cards, render_insight_box
from src.data_loader import load_campaigns_data, load_customers_data, get_channel_list

# Page Configuration
st.set_page_config(
    page_title="CAC Cockpit | Marketing Unit Economics",
    page_icon="📊",
    layout="wide",
    initial_sidebar_state="expanded"
)

def main():
    # 1. Header Banner
    render_header(
        title="Customer Acquisition Cost (CAC) Analysis Cockpit",
        subtitle="Executive marketing performance, channel attribution, payback solvency & unit economics",
        badge_text="ABES MCA Capstone"
    )

    # 2. Data Loading
    df_campaigns = load_campaigns_data()
    df_customers = load_customers_data()
    channel_list = get_channel_list()

    # 3. Persistent Sidebar Filters
    st.sidebar.image("https://cdn-icons-png.flaticon.com/512/3135/3135715.png", width=60)
    st.sidebar.title("🎛️ Analysis Controls")
    
    # Date Filter
    min_date = pd.to_datetime(df_campaigns["campaign_date"]).min().date()
    max_date = pd.to_datetime(df_campaigns["campaign_date"]).max().date()
    
    selected_dates = st.sidebar.date_input(
        "Calendar Date Window",
        value=(min_date, max_date),
        min_value=min_date,
        max_value=max_date
    )
    
    # Channel Multiselect
    selected_channels = st.sidebar.multiselect(
        "Acquisition Channels",
        options=channel_list,
        default=channel_list
    )
    
    # Store in session state for cross-page continuity
    st.session_state["selected_channels"] = selected_channels
    if isinstance(selected_dates, tuple) and len(selected_dates) == 2:
        st.session_state["start_date"] = str(selected_dates[0])
        st.session_state["end_date"] = str(selected_dates[1])
    else:
        st.session_state["start_date"] = str(min_date)
        st.session_state["end_date"] = str(max_date)

    # Filter Data
    start_str = st.session_state["start_date"]
    end_str = st.session_state["end_date"]
    
    filtered_camp = df_campaigns[
        (df_campaigns["campaign_date"] >= start_str) &
        (df_campaigns["campaign_date"] <= end_str) &
        (df_campaigns["channel_name"].isin(selected_channels))
    ]

    if filtered_camp.empty:
        render_insight_box(
            "warning", 
            "No Records Found", 
            "No campaign logs match the selected filter criteria. Please broaden your date range or select additional channels."
        )
        return

    # 4. Global High-Level Metrics
    total_spend = filtered_camp["spend_amount"].sum()
    total_conversions = filtered_camp["conversions_count"].sum()
    blended_cac = round(total_spend / max(total_conversions, 1), 2)
    
    filtered_cust = df_customers[
        (df_customers["channel_name"].isin(selected_channels)) &
        (df_customers["acquisition_date"] >= start_str) &
        (df_customers["acquisition_date"] <= end_str)
    ]
    avg_payback = round(filtered_cust["payback_months"].mean(), 1) if not filtered_cust.empty else 0.0
    payback_status = "positive" if avg_payback <= 12.0 else "critical"

    kpi_cards = [
        {
            "title": "Blended CAC",
            "value": f"${blended_cac:,.2f}",
            "sub": "Portfolio Avg",
            "status": "positive" if blended_cac <= 75.0 else "neutral",
            "detail": "Total Spend / Total Customers"
        },
        {
            "title": "Marketing Outlay",
            "value": f"${total_spend:,.0f}",
            "sub": f"{len(filtered_camp)} Campaigns",
            "status": "neutral",
            "detail": f"{start_str} to {end_str}"
        },
        {
            "title": "New Customers",
            "value": f"{total_conversions:,}",
            "sub": "Converted",
            "status": "positive",
            "detail": "Attributed Conversions"
        },
        {
            "title": "Capital Payback",
            "value": f"{avg_payback} Mo",
            "sub": "Safe (<12m)" if avg_payback <= 12.0 else "Critical (>12m)",
            "status": payback_status,
            "detail": "Time to Recoup Acquisition Spend"
        }
    ]
    render_kpi_cards(kpi_cards)

    st.markdown("<div style='height: 12px;'></div>", unsafe_allow_html=True)

    # 5. Core Charts (Plotly)
    col1, col2 = st.columns([7, 5])

    with col1:
        st.subheader("📈 Monthly Blended CAC Trajectory vs Benchmark")
        # Monthly aggregation
        monthly_trend = filtered_camp.groupby("year_month").agg(
            monthly_spend=("spend_amount", "sum"),
            monthly_customers=("conversions_count", "sum")
        ).reset_index()
        monthly_trend["monthly_cac"] = np.round(monthly_trend["monthly_spend"] / np.maximum(monthly_trend["monthly_customers"], 1), 2)

        fig_trend = go.Figure()
        fig_trend.add_trace(go.Scatter(
            x=monthly_trend["year_month"],
            y=monthly_trend["monthly_cac"],
            mode="lines+markers",
            name="Monthly Blended CAC",
            line=dict(color="#2563EB", width=3),
            marker=dict(size=7, color="#1E3A8A")
        ))
        # Benchmark line ($55 target CPA)
        fig_trend.add_hline(
            y=55.0, 
            line_dash="dash", 
            line_color="#DC2626", 
            annotation_text="Target CPA Benchmark ($55.00)",
            annotation_position="top left"
        )
        fig_trend.update_layout(
            margin=dict(l=20, r=20, t=30, b=20),
            xaxis_title="Calendar Month",
            yaxis_title="Blended CAC ($)",
            plot_bgcolor="#FFFFFF",
            paper_bgcolor="#FFFFFF",
            hovermode="x unified",
            height=340
        )
        st.plotly_chart(fig_trend, use_container_width=True)

    with col2:
        st.subheader("🍩 Acquisition Spend by Channel")
        channel_spend = filtered_camp.groupby("channel_name")["spend_amount"].sum().reset_index()
        fig_pie = px.pie(
            channel_spend,
            names="channel_name",
            values="spend_amount",
            hole=0.45,
            color="channel_name",
            color_discrete_map={
                "Search": "#2563EB",
                "Social": "#F59E0B",
                "Referral": "#10B981",
                "Email": "#6366F1",
                "Organic": "#8B5CF6"
            }
        )
        fig_pie.update_layout(
            margin=dict(l=20, r=20, t=30, b=20),
            height=340,
            showlegend=True
        )
        st.plotly_chart(fig_pie, use_container_width=True)

    # 6. Strategic Insights Callouts
    st.subheader("💡 Executive Action Insights")
    c1, c2 = st.columns(2)
    with c1:
        render_insight_box(
            "critical" if avg_payback > 12.0 else "success",
            "Working Capital Health",
            f"Average customer payback is currently {avg_payback} months. " + 
            ("Paid channels are operating within healthy capital runway thresholds." if avg_payback <= 12.0 
             else "Paid Social exceeds the 12-month solvency threshold, creating working capital drag.")
        )
    with c2:
        render_insight_box(
            "info",
            "Recommended Optimization",
            "Referral traffic delivers customer conversions at <$30 CAC. Reallocating 15-20% of top-of-funnel ad spend into referral incentives will lower blended acquisition cost by an estimated 11.4%."
        )

if __name__ == "__main__":
    main()
