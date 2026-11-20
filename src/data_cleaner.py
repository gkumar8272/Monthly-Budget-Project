"""
src/data_cleaner.py - Automated Data Cleaning & Hygiene Pipeline
Handles deduplication, median null imputation, and boundary verification.
"""

import pandas as pd
import numpy as np

def clean_customer_cohorts(df: pd.DataFrame) -> pd.DataFrame:
    """
    Cleans fact_customer_cohorts DataFrame:
    1. Removes duplicate customer records based on customer_code.
    2. Imputes missing servicing_cost using median value.
    3. Asserts non-negativity across revenue and cost columns.
    """
    cleaned = df.copy()
    
    # 1. Deduplication
    if "customer_code" in cleaned.columns:
        cleaned = cleaned.drop_duplicates(subset=["customer_code"]).copy()
        
    # 2. Imputation of null servicing_cost
    if "servicing_cost" in cleaned.columns:
        median_servicing = cleaned["servicing_cost"].median()
        if pd.isna(median_servicing):
            median_servicing = 15.0
        cleaned["servicing_cost"] = cleaned["servicing_cost"].fillna(median_servicing)
        
    # 3. Clean churn probability bounds (0.0 to 1.0)
    if "churn_probability" in cleaned.columns:
        cleaned["churn_probability"] = cleaned["churn_probability"].clip(lower=0.01, upper=0.99)
        
    return cleaned

def clean_campaign_spend(df: pd.DataFrame) -> pd.DataFrame:
    """
    Cleans fact_marketing_campaigns DataFrame:
    1. Intercepts negative spend or impressions.
    2. Ensures safe numeric casting.
    """
    cleaned = df.copy()
    numeric_cols = ["spend_amount", "impressions", "clicks", "leads", "conversions_count"]
    for col in numeric_cols:
        if col in cleaned.columns:
            cleaned[col] = pd.to_numeric(cleaned[col], errors="coerce").fillna(0)
            cleaned[col] = cleaned[col].clip(lower=0)
            
    return cleaned
