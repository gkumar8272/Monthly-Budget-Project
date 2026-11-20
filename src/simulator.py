"""
src/simulator.py - Budget Reallocation & Marginal ROAS Simulator
Simulates ad spend reallocation across channels on logarithmic diminishing return curves.
"""

import numpy as np
import pandas as pd

def simulate_budget_reallocation(
    campaigns_df: pd.DataFrame, 
    source_channel: str, 
    target_channel: str, 
    reallocation_pct: float
) -> dict:
    """
    Simulates shifting reallocation_pct (e.g. 0.20 = 20%) of source_channel spend into target_channel.
    Uses diminishing returns curves:
    - Target channel customer conversion follows diminishing elasticity exponent (e.g. 0.82)
    Returns:
        dict containing:
        - baseline_spend
        - baseline_customers
        - baseline_blended_cac
        - simulated_spend
        - simulated_customers
        - simulated_blended_cac
        - customer_gain
        - blended_cac_change_pct
        - channel_comparison_df
    """
    if campaigns_df.empty:
        return {}
        
    summary = campaigns_df.groupby("channel_name").agg(
        spend=("spend_amount", "sum"),
        customers=("conversions_count", "sum")
    ).reset_index()
    
    baseline_spend = float(summary["spend"].sum())
    baseline_customers = int(summary["customers"].sum())
    baseline_blended_cac = round(baseline_spend / max(baseline_customers, 1), 2)
    
    # Calculate shifted amount
    source_row = summary[summary["channel_name"] == source_channel]
    if source_row.empty:
        return {}
        
    source_spend = float(source_row["spend"].iloc[0])
    shift_amount = source_spend * (reallocation_pct / 100.0)
    
    # Calculate simulated changes
    sim_data = []
    total_sim_customers = 0
    total_sim_spend = 0.0
    
    for _, row in summary.iterrows():
        ch = row["channel_name"]
        curr_spend = float(row["spend"])
        curr_cust = int(row["customers"])
        curr_cac = round(curr_spend / max(curr_cust, 1), 2)
        
        if ch == source_channel:
            new_spend = max(curr_spend - shift_amount, 0.0)
            # Scaling down customers linearly or slightly sub-linear
            ratio = new_spend / max(curr_spend, 1.0)
            new_cust = int(round(curr_cust * (ratio ** 0.95)))
        elif ch == target_channel:
            new_spend = curr_spend + shift_amount
            # Diminishing returns scaling: customers scale by (new_spend / curr_spend)^0.82
            ratio = new_spend / max(curr_spend, 1.0)
            new_cust = int(round(curr_cust * (ratio ** 0.82)))
        else:
            new_spend = curr_spend
            new_cust = curr_cust
            
        new_cac = round(new_spend / max(new_cust, 1), 2)
        total_sim_spend += new_spend
        total_sim_customers += new_cust
        
        sim_data.append({
            "channel_name": ch,
            "baseline_spend": round(curr_spend, 2),
            "simulated_spend": round(new_spend, 2),
            "baseline_customers": curr_cust,
            "simulated_customers": new_cust,
            "baseline_cac": curr_cac,
            "simulated_cac": new_cac
        })
        
    sim_blended_cac = round(total_sim_spend / max(total_sim_customers, 1), 2)
    customer_gain = total_sim_customers - baseline_customers
    cac_change_pct = round(((sim_blended_cac - baseline_blended_cac) / baseline_blended_cac) * 100.0, 2)
    
    return {
        "baseline_spend": round(baseline_spend, 2),
        "baseline_customers": baseline_customers,
        "baseline_blended_cac": baseline_blended_cac,
        "simulated_spend": round(total_sim_spend, 2),
        "simulated_customers": total_sim_customers,
        "simulated_blended_cac": sim_blended_cac,
        "customer_gain": customer_gain,
        "cac_change_pct": cac_change_pct,
        "comparison_df": pd.DataFrame(sim_data)
    }
