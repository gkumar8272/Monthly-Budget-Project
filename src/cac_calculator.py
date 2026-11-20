"""
src/cac_calculator.py - Business Logic & Unit Economics Computation Engine
Pure functional calculations for Blended CAC, Channel CAC, Payback Period, and LTV.
"""

import numpy as np
import pandas as pd

def calculate_blended_cac(total_spend: float, total_conversions: int) -> float:
    """Calculates overall blended CAC with division guard."""
    if total_conversions <= 0:
        return 0.0
    return round(float(total_spend) / float(total_conversions), 2)

def calculate_channel_cac_summary(campaigns_df: pd.DataFrame) -> pd.DataFrame:
    """
    Computes aggregated CAC, total spend, and conversions grouped by channel.
    """
    if campaigns_df.empty:
        return pd.DataFrame(columns=["channel_name", "total_spend", "total_conversions", "channel_cac"])
        
    grouped = campaigns_df.groupby("channel_name").agg(
        total_spend=("spend_amount", "sum"),
        total_impressions=("impressions", "sum"),
        total_clicks=("clicks", "sum"),
        total_leads=("leads", "sum"),
        total_conversions=("conversions_count", "sum")
    ).reset_index()
    
    # Safe division
    grouped["channel_cac"] = np.where(
        grouped["total_conversions"] > 0,
        np.round(grouped["total_spend"] / grouped["total_conversions"], 2),
        0.0
    )
    grouped["ctr_pct"] = np.where(
        grouped["total_impressions"] > 0,
        np.round((grouped["total_clicks"] / grouped["total_impressions"]) * 100.0, 2),
        0.0
    )
    grouped["lead_cvr_pct"] = np.where(
        grouped["total_leads"] > 0,
        np.round((grouped["total_conversions"] / grouped["total_leads"]) * 100.0, 2),
        0.0
    )
    grouped["cpc"] = np.where(
        grouped["total_clicks"] > 0,
        np.round(grouped["total_spend"] / grouped["total_clicks"], 2),
        0.0
    )
    grouped["cpl"] = np.where(
        grouped["total_leads"] > 0,
        np.round(grouped["total_spend"] / grouped["total_leads"], 2),
        0.0
    )
    
    return grouped

def calculate_payback_months(cac: float, mrr: float, gross_margin_pct: float) -> float:
    """Calculates months required to recoup CAC."""
    monthly_gp = mrr * gross_margin_pct
    if monthly_gp <= 0:
        return 999.0 # Infinity representation
    return round(cac / monthly_gp, 1)

def compute_tier_unit_economics(customers_df: pd.DataFrame) -> pd.DataFrame:
    """
    Aggregates LTV:CAC, average tenure, MRR, and payback by subscription tier.
    """
    if customers_df.empty:
        return pd.DataFrame()
        
    tier_summary = customers_df.groupby("tier_name").agg(
        total_customers=("customer_id", "count"),
        avg_mrr=("monthly_revenue", "mean"),
        avg_cac=("cac_attributed", "mean"),
        avg_tenure=("tenure_months", "mean"),
        avg_ltv=("lifetime_value", "mean"),
        avg_payback=("payback_months", "mean")
    ).reset_index()
    
    tier_summary["ltv_cac_ratio"] = np.where(
        tier_summary["avg_cac"] > 0,
        np.round(tier_summary["avg_ltv"] / tier_summary["avg_cac"], 2),
        0.0
    )
    
    # Classify health status
    conditions = [
        (tier_summary["ltv_cac_ratio"] >= 3.0),
        (tier_summary["ltv_cac_ratio"] >= 1.0) & (tier_summary["ltv_cac_ratio"] < 3.0),
        (tier_summary["ltv_cac_ratio"] < 1.0)
    ]
    choices = ["Healthy (>= 3.0x)", "Sub-Optimal (1.0 - 2.9x)", "Value-Destroying (< 1.0x)"]
    tier_summary["health_category"] = np.select(conditions, choices, default="Unknown")
    
    return tier_summary

def recompute_churn_sensitivity(customers_df: pd.DataFrame, churn_multiplier: float) -> pd.DataFrame:
    """
    Re-estimates customer LTV and LTV:CAC when user adjusts baseline churn.
    churn_multiplier: 1.0 = baseline, 1.25 = +25% churn, 0.8 = -20% churn
    """
    df = customers_df.copy()
    # Baseline expected lifespan = 1 / churn_rate
    # If churn increases by multiplier, lifespan scales inversely: adjusted_tenure = tenure / multiplier
    adjusted_tenure = np.maximum(df["tenure_months"] / max(churn_multiplier, 0.1), 1.0)
    df["adjusted_ltv"] = np.round(df["monthly_revenue"] * adjusted_tenure * df["gross_margin_pct"], 2)
    df["adjusted_ltv_cac"] = np.where(
        df["cac_attributed"] > 0,
        np.round(df["adjusted_ltv"] / df["cac_attributed"], 2),
        0.0
    )
    return df
