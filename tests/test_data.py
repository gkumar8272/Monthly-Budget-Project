"""
tests/test_data.py - Unit & Data Integrity Test Suite
TC-T-01 to TC-T-05: Database persistence, schema structure, PK/FK constraints, and cleaning logic.
"""

import os
import sqlite3
import pytest
import pandas as pd
import numpy as np

DB_PATH = os.path.abspath("data/app.db")

@pytest.fixture(scope="module")
def db_conn():
    """Provides a read-only connection to data/app.db for testing."""
    assert os.path.exists(DB_PATH), f"Database file not found at {DB_PATH}"
    conn = sqlite3.connect(f"file:{DB_PATH}?mode=ro", uri=True)
    yield conn
    conn.close()

def test_tc_t_01_database_connectivity_and_tables(db_conn):
    """TC-T-01: Verify SQLite connection and existence of all 4 tables + 1 view."""
    cur = db_conn.cursor()
    cur.execute("SELECT name, type FROM sqlite_master WHERE type IN ('table', 'view');")
    entities = {row[0]: row[1] for row in cur.fetchall()}
    
    assert "dim_channels" in entities, "Table dim_channels missing"
    assert "dim_subscription_tiers" in entities, "Table dim_subscription_tiers missing"
    assert "fact_marketing_campaigns" in entities, "Table fact_marketing_campaigns missing"
    assert "fact_customer_cohorts" in entities, "Table fact_customer_cohorts missing"
    assert "view_cac_full_analysis" in entities, "View view_cac_full_analysis missing"

def test_tc_t_02_table_row_and_column_counts(db_conn):
    """TC-T-02: Verify expected row counts and column schemas."""
    cur = db_conn.cursor()
    
    cur.execute("SELECT COUNT(*) FROM dim_channels;")
    assert cur.fetchone()[0] == 5, "dim_channels must contain exactly 5 channels"
    
    cur.execute("SELECT COUNT(*) FROM dim_subscription_tiers;")
    assert cur.fetchone()[0] == 4, "dim_subscription_tiers must contain exactly 4 tiers"
    
    cur.execute("SELECT COUNT(*) FROM fact_marketing_campaigns;")
    assert cur.fetchone()[0] == 240, "fact_marketing_campaigns must contain 240 monthly campaign records"
    
    cur.execute("SELECT COUNT(*) FROM fact_customer_cohorts;")
    assert cur.fetchone()[0] >= 300, "fact_customer_cohorts must contain at least 300 customer records"
    
    cur.execute("PRAGMA table_info(view_cac_full_analysis);")
    cols = cur.fetchall()
    assert len(cols) == 34, f"view_cac_full_analysis must possess exactly 34 columns, found {len(cols)}"

def test_tc_t_03_primary_key_uniqueness(db_conn):
    """TC-T-03: Verify uniqueness of primary keys across master and fact tables."""
    cur = db_conn.cursor()
    
    cur.execute("SELECT channel_id, COUNT(*) FROM dim_channels GROUP BY channel_id HAVING COUNT(*) > 1;")
    assert len(cur.fetchall()) == 0, "Duplicate channel_id found in dim_channels"
    
    cur.execute("SELECT tier_id, COUNT(*) FROM dim_subscription_tiers GROUP BY tier_id HAVING COUNT(*) > 1;")
    assert len(cur.fetchall()) == 0, "Duplicate tier_id found in dim_subscription_tiers"
    
    cur.execute("SELECT campaign_id, COUNT(*) FROM fact_marketing_campaigns GROUP BY campaign_id HAVING COUNT(*) > 1;")
    assert len(cur.fetchall()) == 0, "Duplicate campaign_id found in fact_marketing_campaigns"

def test_tc_t_04_foreign_key_referential_integrity(db_conn):
    """TC-T-04: Execute PRAGMA foreign_key_check to verify zero broken foreign keys."""
    cur = db_conn.cursor()
    cur.execute("PRAGMA foreign_key_check;")
    fk_violations = cur.fetchall()
    assert len(fk_violations) == 0, f"Found {len(fk_violations)} FK violations: {fk_violations}"

def test_tc_t_05_data_cleaning_pipeline():
    """TC-T-05: Test automated deduplication and median null imputation logic on customer data."""
    # Mock messy dataset containing duplicates and nulls
    raw_df = pd.DataFrame({
        "customer_id": [1, 2, 3, 4, 4], # duplicate ID 4
        "customer_code": ["C-1", "C-2", "C-3", "C-4", "C-4"],
        "servicing_cost": [10.0, np.nan, 30.0, 40.0, 40.0] # 1 null
    })
    
    # Cleaning step 1: Deduplication
    clean_df = raw_df.drop_duplicates(subset=["customer_code"]).copy()
    assert len(clean_df) == 4, "Deduplication must eliminate duplicate row"
    
    # Cleaning step 2: Median Imputation
    median_val = clean_df["servicing_cost"].median()
    clean_df["servicing_cost"] = clean_df["servicing_cost"].fillna(median_val)
    assert clean_df["servicing_cost"].isnull().sum() == 0, "Null servicing_cost must be imputed"
    assert clean_df.loc[clean_df["customer_id"] == 2, "servicing_cost"].iloc[0] == median_val
