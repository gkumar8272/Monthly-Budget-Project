"""
pages/5_Budget_Simulator.py - Strategic Unit Economics & Budget Sandbox (UI-06)
"""

import streamlit as st
import pandas as pd
import numpy as np
import plotly.express as px
import plotly.graph_objects as go

from src.ui import render_header, render_kpi_cards, render_insight_box
from src.data_loader import load_campaigns_data, load_customers_data
from src.cac_calculator import compute_tier_unit_economics, recompute_churn_sensitivity
from src.simulator import simulate_budget_reallocation

st.set_page_config(page_title="Budget Simulator & Unit Economics", page_icon="🧮", layout="wide")

def main():
    render_header(
        title="Strategic Unit Economics & Budget Reallocation Sandbox",
        subtitle="Simulate marginal ROAS returns, tiered LTV:CAC health, and macroeconomic churn sensitivity",
        badge_text="UI-06 Simulation Engine"
    )

    df_camp = load_campaigns_data()
    df_cust = load_customers_data()

    tabs = st.tabs(["📊 Tiered LTV:CAC Matrix", "🔀 Budget Reallocation Sandbox", "📉 Churn Sensitivity Modeling"])

    # TAB 1: Tiered LTV:CAC
    with tabs[0]:
        st.subheader("Customer Lifetime Value to CAC (LTV:CAC) by Subscription Tier")
        tier_df = compute_tier_unit_economics(df_cust)

        col1, col2 = st.columns([7, 5])
        with col1:
            fig_ltv = go.Figure()
            fig_ltv.add_trace(go.Bar(
                x=tier_df["tier_name"],
                y=tier_df["ltv_cac_ratio"],
                text=tier_df["ltv_cac_ratio"].apply(lambda v: f"{v:.1f}x"),
                textposition="outside",
                marker_color=["#EF4444" if r < 3.0 else "#10B981" for r in tier_df["ltv_cac_ratio"]]
            ))
            fig_ltv.add_hline(y=3.0, line_dash="dash", line_color="#DC2626", annotation_text="Viability Threshold (3.0x)")
            fig_ltv.update_layout(
                xaxis_title="Subscription Plan Category",
                yaxis_title="LTV : CAC Ratio",
                plot_bgcolor="#FFFFFF",
                paper_bgcolor="#FFFFFF",
                height=350
            )
            st.plotly_chart(fig_ltv, use_container_width=True)

        with col2:
            st.dataframe(
                tier_df[["tier_name", "total_customers", "avg_mrr", "avg_cac", "avg_ltv", "ltv_cac_ratio", "health_category"]].rename(columns={
                    "tier_name": "Tier",
                    "total_customers": "Accounts",
                    "avg_mrr": "MRR ($)",
                    "avg_cac": "CAC ($)",
                    "avg_ltv": "LTV ($)",
                    "ltv_cac_ratio": "Ratio",
                    "health_category": "Health Status"
                }),
                use_container_width=True,
                hide_index=True
            )
            render_insight_box(
                "warning",
                "Tier Viability Finding",
                "Basic plan tier operates at an LTV:CAC ratio below 3.0x due to high annual logo churn (24%). Enterprise and SMB tiers exhibit outstanding unit economics (> 5.0x ratio)."
            )

    # TAB 2: Budget Reallocation Sandbox
    with tabs[1]:
        st.subheader("Simulate Marketing Budget Shifts & Diminishing Returns")
        c1, c2, c3 = st.columns(3)
        with c1:
            src_ch = st.selectbox("Shift Budget Out Of (Source)", options=["Social", "Search", "Email"], index=0)
        with c2:
            tgt_ch = st.selectbox("Reallocate Budget Into (Target)", options=["Referral", "Search", "Organic"], index=0)
        with c3:
            shift_pct = st.slider("Budget Shift Percentage (%)", min_value=5, max_value=50, value=20, step=5)

        sim_res = simulate_budget_reallocation(df_camp, src_ch, tgt_ch, shift_pct)

        if sim_res:
            render_kpi_cards([
                {"title": "Baseline Customers", "value": f"{sim_res['baseline_customers']:,}", "sub": "Current State", "status": "neutral"},
                {"title": "Simulated Customers", "value": f"{sim_res['simulated_customers']:,}", "sub": f"+{sim_res['customer_gain']} Gain", "status": "positive"},
                {"title": "Baseline Blended CAC", "value": f"${sim_res['baseline_blended_cac']:.2f}", "sub": "Current CAC", "status": "neutral"},
                {"title": "Simulated Blended CAC", "value": f"${sim_res['simulated_blended_cac']:.2f}", "sub": f"{sim_res['cac_change_pct']}% Delta", "status": "positive"}
            ])

            st.dataframe(
                sim_res["comparison_df"].rename(columns={
                    "channel_name": "Channel",
                    "baseline_spend": "Current Spend ($)",
                    "simulated_spend": "New Spend ($)",
                    "baseline_customers": "Current Cust",
                    "simulated_customers": "New Cust",
                    "baseline_cac": "Current CAC ($)",
                    "simulated_cac": "New CAC ($)"
                }),
                use_container_width=True,
                hide_index=True
            )

    # TAB 3: Churn Sensitivity
    with tabs[2]:
        st.subheader("Macroeconomic Churn Sensitivity Stress-Test")
        churn_slider = st.slider("Macro Churn Multiplier (-50% to +50%)", min_value=-50, max_value=50, value=15, step=5)
        multiplier = 1.0 + (churn_slider / 100.0)
        
        sens_df = recompute_churn_sensitivity(df_cust, multiplier)
        
        base_ltv = df_cust["lifetime_value"].mean()
        new_ltv = sens_df["adjusted_ltv"].mean()
        ltv_drop_pct = round(((new_ltv - base_ltv) / base_ltv) * 100.0, 1)

        st.metric(
            label="Portfolio Average LTV Stress Response",
            value=f"${new_ltv:,.2f}",
            delta=f"{ltv_drop_pct}% vs Baseline (${base_ltv:,.2f})"
        )

        render_insight_box(
            "info",
            "Sensitivity Finding",
            f"Under a {churn_slider}% macroeconomic churn shift, expected customer lifetime gross margin decreases by {abs(ltv_drop_pct)}%. Enterprise tiers remain resilient while Basic tiers experience severe capital deficit."
        )

if __name__ == "__main__":
    main()
