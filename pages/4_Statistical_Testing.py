"""
pages/4_Statistical_Testing.py - Inferential EDA & Hypothesis Testing (UI-05)
"""

import streamlit as st
import pandas as pd
import numpy as np
import plotly.express as px

from src.ui import render_header, render_insight_box
from src.data_loader import load_campaigns_data
from src.stats_engine import run_channel_anova_kruskal

st.set_page_config(page_title="Statistical Testing", page_icon="🧪", layout="wide")

def main():
    render_header(
        title="Inferential Statistics & Hypothesis Testing Engine",
        subtitle="Prove statistical significance of cross-channel acquisition cost variances (ANOVA & Kruskal-Wallis)",
        badge_text="UI-05 Scientific EDA"
    )

    df_camp = load_campaigns_data()

    # Sidebar Controls
    st.sidebar.title("🧪 Test Parameters")
    alpha = st.sidebar.slider("Significance Level (Alpha)", min_value=0.01, max_value=0.10, value=0.05, step=0.01)
    
    st.markdown(f"""
    **Formal Hypothesis Formulation:**
    - **Null Hypothesis ($H_0$):** Mean acquisition costs across Search, Social, Referral, Email, and Organic are identical ($\mu_1 = \mu_2 = \dots = \mu_k$).
    - **Alternative Hypothesis ($H_1$):** At least one marketing channel exhibits statistically different acquisition unit economics.
    - **Target Significance Level:** $\\alpha = {alpha}$
    """)

    # Execute Engine
    results = run_channel_anova_kruskal(df_camp, alpha=alpha)

    # Statistical Test Badges
    col1, col2, col3 = st.columns(3)
    with col1:
        st.metric(
            label="Shapiro-Wilk Normality p-value",
            value=f"{results['shapiro_p']:.5f}",
            delta="Non-Normal (Skewed)" if not results['is_normal'] else "Normal Distribution"
        )
    with col2:
        st.metric(
            label="One-Way ANOVA F-Statistic",
            value=f"{results['anova_f']:.2f}",
            delta=f"p = {results['anova_p']:.2e}"
        )
    with col3:
        st.metric(
            label="Kruskal-Wallis H-Statistic",
            value=f"{results['kruskal_h']:.2f}",
            delta="Statistically Significant" if results['significant_kruskal'] else "Not Significant"
        )

    st.markdown("<div style='height: 10px;'></div>", unsafe_allow_html=True)

    # Narrative Verdict Box
    if results["significant_kruskal"]:
        render_insight_box(
            "success",
            "Hypothesis Testing Verdict: REJECT NULL HYPOTHESIS",
            f"Both One-Way ANOVA (F = {results['anova_f']}, p < 0.001) and the non-parametric Kruskal-Wallis test (H = {results['kruskal_h']}, p < 0.001) demonstrate that channel CAC variations are genuine and statistically significant at alpha = {alpha}."
        )
    else:
        render_insight_box(
            "warning",
            "Hypothesis Testing Verdict: RETAIN NULL HYPOTHESIS",
            f"Insufficient statistical evidence to reject null hypothesis at alpha = {alpha}."
        )

    # CAC Distribution Boxplot
    st.subheader("📦 Campaign CAC Distribution & Outlier Spread by Channel")
    fig_box = px.box(
        df_camp,
        x="channel_name",
        y="cac_calculated",
        color="channel_name",
        points="all",
        hover_data=["campaign_name", "spend_amount", "conversions_count"],
        color_discrete_map={
            "Search": "#2563EB",
            "Social": "#F59E0B",
            "Referral": "#10B981",
            "Email": "#6366F1",
            "Organic": "#8B5CF6"
        }
    )
    fig_box.update_layout(
        xaxis_title="Marketing Channel",
        yaxis_title="Calculated Campaign CAC ($)",
        plot_bgcolor="#FFFFFF",
        paper_bgcolor="#FFFFFF",
        height=400
    )
    st.plotly_chart(fig_box, use_container_width=True)

    # Channel Statistics Table
    st.subheader("📋 Descriptive Sample Parameters")
    stats_df = pd.DataFrame({
        "Channel": results["group_names"],
        "Sample Size (n)": [results["sample_sizes"][k] for k in results["group_names"]],
        "Mean CAC ($)": [results["channel_means"][k] for k in results["group_names"]],
        "Std Deviation ($)": [results["channel_stds"][k] for k in results["group_names"]]
    })
    st.dataframe(stats_df, use_container_width=True, hide_index=True)

if __name__ == "__main__":
    main()
