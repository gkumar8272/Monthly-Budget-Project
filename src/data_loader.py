"""
src/data_loader.py - High-Performance Cached Data Extraction
Wraps SQLite queries with Streamlit @st.cache_data for instant in-memory retrieval.
"""

import sqlite3
import pandas as pd
import streamlit as st
from src.db import get_db_connection
from src.data_cleaner import clean_customer_cohorts, clean_campaign_spend

@st.cache_data(ttl=3600, show_spinner=False)
def load_full_analysis_data() -> pd.DataFrame:
    """
    Loads and cleans the 34-column pre-joined analytical view: view_cac_full_analysis.
    """
    conn = get_db_connection(read_only=True)
    df = pd.read_sql_query("SELECT * FROM view_cac_full_analysis;", conn)
    conn.close()
    
    # Deduplicate by customer_code if present
    if "customer_code" in df.columns:
        df = df.drop_duplicates(subset=["customer_code"]).copy()
        
    return df

@st.cache_data(ttl=3600, show_spinner=False)
def load_campaigns_data() -> pd.DataFrame:
    """
    Loads fact_marketing_campaigns joined with dim_channels.
    """
    conn = get_db_connection(read_only=True)
    query = """
    SELECT 
        c.campaign_id,
        c.channel_id,
        ch.channel_name,
        ch.channel_type,
        c.campaign_name,
        c.campaign_date,
        c.year_month,
        c.seasonal_period,
        c.spend_amount,
        c.impressions,
        c.clicks,
        c.leads,
        c.conversions_count,
        c.cac_calculated,
        c.cpc_calculated,
        c.cpl_calculated
    FROM fact_marketing_campaigns c
    JOIN dim_channels ch ON c.channel_id = ch.channel_id
    ORDER BY c.campaign_date ASC;
    """
    df = pd.read_sql_query(query, conn)
    conn.close()
    return clean_campaign_spend(df)

@st.cache_data(ttl=3600, show_spinner=False)
def load_customers_data() -> pd.DataFrame:
    """
    Loads fact_customer_cohorts joined with dimensions.
    """
    conn = get_db_connection(read_only=True)
    query = """
    SELECT 
        cu.*,
        ch.channel_name,
        tr.tier_name,
        tr.gross_margin_pct
    FROM fact_customer_cohorts cu
    JOIN dim_channels ch ON cu.channel_id = ch.channel_id
    JOIN dim_subscription_tiers tr ON cu.tier_id = tr.tier_id;
    """
    df = pd.read_sql_query(query, conn)
    conn.close()
    return clean_customer_cohorts(df)

@st.cache_data(ttl=3600, show_spinner=False)
def get_channel_list() -> list:
    """Returns list of active channel names."""
    conn = get_db_connection(read_only=True)
    cur = conn.cursor()
    cur.execute("SELECT channel_name FROM dim_channels WHERE is_active = 1 ORDER BY channel_id;")
    channels = [row[0] for row in cur.fetchall()]
    conn.close()
    return channels

@st.cache_data(ttl=3600, show_spinner=False)
def get_tier_list() -> list:
    """Returns list of subscription tier names."""
    conn = get_db_connection(read_only=True)
    cur = conn.cursor()
    cur.execute("SELECT tier_name FROM dim_subscription_tiers ORDER BY tier_id;")
    tiers = [row[0] for row in cur.fetchall()]
    conn.close()
    return tiers
