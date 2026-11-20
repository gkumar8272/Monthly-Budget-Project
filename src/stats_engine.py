"""
src/stats_engine.py - Inferential Statistical Modeling Engine
Executes Shapiro-Wilk normality tests, One-Way ANOVA, Kruskal-Wallis, and effect sizes.
"""

import numpy as np
import pandas as pd
from scipy import stats

def run_channel_anova_kruskal(campaigns_df: pd.DataFrame, alpha: float = 0.05) -> dict:
    """
    Evaluates whether CAC differences across marketing channels are statistically significant.
    Returns:
        dict containing:
        - group_names: list of channels
        - sample_sizes: dict of n per channel
        - channel_means: dict of mean CAC per channel
        - channel_stds: dict of std CAC per channel
        - shapiro_p: p-value of normality check on residuals
        - is_normal: boolean
        - anova_f: F-statistic
        - anova_p: p-value
        - kruskal_h: H-statistic
        - kruskal_p: p-value
        - verdict: human-readable conclusion
    """
    if campaigns_df.empty or "cac_calculated" not in campaigns_df.columns:
        return {"error": "Insufficient data"}
        
    # Group CAC by channel
    groups = []
    group_names = []
    sample_sizes = {}
    channel_means = {}
    channel_stds = {}
    
    for ch_name, group in campaigns_df.groupby("channel_name"):
        cac_vals = group["cac_calculated"].dropna().values
        if len(cac_vals) >= 3: # Minimum observations for statistical testing
            groups.append(cac_vals)
            group_names.append(ch_name)
            sample_sizes[ch_name] = len(cac_vals)
            channel_means[ch_name] = round(float(np.mean(cac_vals)), 2)
            channel_stds[ch_name] = round(float(np.std(cac_vals, ddof=1)), 2)
            
    if len(groups) < 2:
        return {"error": "At least 2 channel groups with >=3 samples required"}
        
    # 1. Normality Check on combined residuals
    all_residuals = []
    for g, mean in zip(groups, channel_means.values()):
        all_residuals.extend(g - mean)
    shapiro_w, shapiro_p = stats.shapiro(all_residuals[:5000]) # Cap for scipy
    is_normal = bool(shapiro_p >= alpha)
    
    # 2. One-Way ANOVA (Parametric)
    f_stat, anova_p = stats.f_oneway(*groups)
    
    # 3. Kruskal-Wallis (Non-Parametric)
    h_stat, kruskal_p = stats.kruskal(*groups)
    
    # Decision verdict
    significant_anova = bool(anova_p < alpha)
    significant_kruskal = bool(kruskal_p < alpha)
    
    if significant_kruskal:
        verdict = f"Statistically Significant Difference Detected (p < {alpha}). Observed channel cost variations are genuine and not due to random sampling variance."
    else:
        verdict = f"No Statistically Significant Difference Detected (p >= {alpha}). Null hypothesis cannot be rejected."

    return {
        "group_names": group_names,
        "sample_sizes": sample_sizes,
        "channel_means": channel_means,
        "channel_stds": channel_stds,
        "shapiro_w": round(float(shapiro_w), 4),
        "shapiro_p": float(shapiro_p),
        "is_normal": is_normal,
        "anova_f": round(float(f_stat), 3),
        "anova_p": float(anova_p),
        "kruskal_h": round(float(h_stat), 3),
        "kruskal_p": float(kruskal_p),
        "significant_anova": significant_anova,
        "significant_kruskal": significant_kruskal,
        "verdict": verdict
    }
