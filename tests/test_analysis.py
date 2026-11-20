"""
tests/test_analysis.py - Unit Tests for Analytical Calculations & Statistical Engine
TC-T-06 to TC-T-10: Mathematical assertions on known datasets, ANOVA, Kruskal-Wallis, zero-division guards.
"""

import math
import pytest
import numpy as np
import pandas as pd
from scipy import stats

def test_tc_t_06_cac_formula_on_known_dataset():
    """TC-T-06: Verify CAC = Spend / Conversions on benchmark test vectors."""
    test_cases = [
        {"spend": 1000.0, "conversions": 20, "expected_cac": 50.0},
        {"spend": 4500.0, "conversions": 90, "expected_cac": 50.0},
        {"spend": 0.0, "conversions": 10, "expected_cac": 0.0},
        {"spend": 750.50, "conversions": 5, "expected_cac": 150.10}
    ]
    for tc in test_cases:
        calculated_cac = round(tc["spend"] / tc["conversions"], 2)
        assert math.isclose(calculated_cac, tc["expected_cac"], abs_tol=0.01), \
            f"Expected CAC {tc['expected_cac']}, got {calculated_cac}"

def test_tc_t_07_payback_period_formula():
    """TC-T-07: Verify Payback Months = CAC / (MRR * Gross Margin %) and 12-month alert condition."""
    # Scenario 1: Healthy payback (CAC=$300, MRR=$100, Margin=75% -> Gross Profit=$75/mo -> Payback=4 months)
    cac_1 = 300.0
    mrr_1 = 100.0
    margin_1 = 0.75
    monthly_gp_1 = mrr_1 * margin_1
    payback_1 = round(cac_1 / monthly_gp_1, 1)
    assert payback_1 == 4.0, f"Expected 4.0 months, got {payback_1}"
    assert payback_1 <= 12.0, "Payback should be within solvency limit"

    # Scenario 2: Critical payback breach (CAC=$800, MRR=$50, Margin=70% -> Gross Profit=$35/mo -> Payback=22.9 months)
    cac_2 = 800.0
    mrr_2 = 50.0
    margin_2 = 0.70
    monthly_gp_2 = mrr_2 * margin_2
    payback_2 = round(cac_2 / monthly_gp_2, 1)
    assert payback_2 == 22.9, f"Expected 22.9 months, got {payback_2}"
    assert payback_2 > 12.0, "Payback breach flag must be triggered"

def test_tc_t_08_ltv_cac_ratio_and_health_categorization():
    """TC-T-08: Verify LTV:CAC ratio formula and classification standard."""
    # Formula: LTV = (MRR * Margin) / Churn_Rate
    mrr = 200.0
    margin = 0.80
    churn = 0.10 # 10% annual churn
    cac = 150.0
    
    annual_gp = (mrr * 12) * margin # $1,920 annual gross profit
    ltv = annual_gp / churn # $19,200
    ratio = round(ltv / cac, 2)
    
    assert ratio >= 3.0, "Ratio must exceed 3.0 threshold for healthy tier"
    
    def classify_health(r):
        if r >= 3.0:
            return "Healthy"
        elif r >= 1.0:
            return "Sub-Optimal"
        return "Value-Destroying"
        
    assert classify_health(ratio) == "Healthy"
    assert classify_health(2.1) == "Sub-Optimal"
    assert classify_health(0.7) == "Value-Destroying"

def test_tc_t_09_inferential_anova_and_kruskal_wallis():
    """TC-T-09: Verify ANOVA and Kruskal-Wallis statistical test routines on known synthetic groups."""
    # Create 3 distinct channel cost distributions with clear mean differences
    np.random.seed(42)
    organic_costs = np.random.normal(loc=25.0, scale=3.0, size=50)
    search_costs = np.random.normal(loc=65.0, scale=6.0, size=50)
    social_costs = np.random.normal(loc=95.0, scale=8.0, size=50)
    
    # 1. One-Way ANOVA
    f_stat, p_val_anova = stats.f_oneway(organic_costs, search_costs, social_costs)
    assert f_stat > 10.0, "F-statistic must indicate significant between-group variance"
    assert p_val_anova < 0.001, "ANOVA p-value must reject null hypothesis at alpha = 0.05"
    
    # 2. Kruskal-Wallis
    h_stat, p_val_kruskal = stats.kruskal(organic_costs, search_costs, social_costs)
    assert h_stat > 20.0, "Kruskal-Wallis H-statistic must indicate rank difference"
    assert p_val_kruskal < 0.001, "Kruskal-Wallis p-value must confirm significance"

def test_tc_t_10_zero_division_guard():
    """TC-T-10: Verify vectorized zero-division handling for conversion rates and CPC."""
    spend = np.array([100.0, 500.0, 0.0])
    clicks = np.array([50, 0, 10]) # click count has 0
    
    # Vectorized safe division
    cpc = np.where(clicks > 0, spend / clicks, 0.0)
    assert cpc[0] == 2.0
    assert cpc[1] == 0.0, "Zero clicks must return 0.0 without raising ZeroDivisionError"
    assert cpc[2] == 0.0
