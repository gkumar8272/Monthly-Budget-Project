"""
generate_eda_dataset.py — Synthetic CAC Marketing Dataset (550 rows, 18 columns)
Seed 42 | Indian locale | Built-in quality issues for cleaning lessons
"""
import numpy as np
import pandas as pd
from datetime import datetime, timedelta

np.random.seed(42)

# ── Master Dimensions ──────────────────────────────────────────────────────
channels = ['Organic Search', 'Paid Search', 'Paid Social', 'Email', 'Referral']
cities = ['Mumbai', 'Delhi', 'Bengaluru', 'Hyderabad', 'Pune',
          'Chennai', 'Kolkata', 'Ahmedabad', 'Jaipur', 'Lucknow']
tiers = ['Starter', 'Growth', 'Professional', 'Enterprise']
statuses = ['Active', 'Paused', 'Completed']

# Channel cost profiles (mean CAC, std)
channel_profiles = {
    'Organic Search':  {'cac_mean': 420,  'cac_std': 120, 'ctr': 0.045, 'conv_rate': 0.08},
    'Paid Search':     {'cac_mean': 980,  'cac_std': 250, 'ctr': 0.032, 'conv_rate': 0.05},
    'Paid Social':     {'cac_mean': 1400, 'cac_std': 400, 'ctr': 0.018, 'conv_rate': 0.03},
    'Email':           {'cac_mean': 650,  'cac_std': 180, 'ctr': 0.055, 'conv_rate': 0.10},
    'Referral':        {'cac_mean': 350,  'cac_std': 90,  'ctr': 0.060, 'conv_rate': 0.12},
}

# Tier LTV profiles
tier_profiles = {
    'Starter':       {'ltv_mean': 8000,  'ltv_std': 2000, 'rev_mean': 2500,  'churn_mean': 0.22},
    'Growth':        {'ltv_mean': 22000, 'ltv_std': 5000, 'rev_mean': 8000,  'churn_mean': 0.14},
    'Professional':  {'ltv_mean': 45000, 'ltv_std': 8000, 'rev_mean': 18000, 'churn_mean': 0.08},
    'Enterprise':    {'ltv_mean': 72000, 'ltv_std': 12000,'rev_mean': 38000, 'churn_mean': 0.04},
}

# ── Date range: Jan 2023 – Jun 2025 (30 months) ───────────────────────────
dates = pd.date_range('2023-01-01', '2025-06-01', freq='MS')
n_rows = 550

rows = []
for i in range(n_rows):
    dt = np.random.choice(dates)
    ch = np.random.choice(channels, p=[0.22, 0.25, 0.20, 0.18, 0.15])
    city = np.random.choice(cities)
    tier = np.random.choice(tiers, p=[0.30, 0.30, 0.25, 0.15])
    status = np.random.choice(statuses, p=[0.45, 0.15, 0.40])
    
    prof = channel_profiles[ch]
    tier_p = tier_profiles[tier]
    
    # Seasonal Q4 spike for Paid Social
    month = pd.Timestamp(dt).month
    quarter = f"Q{(month - 1) // 3 + 1}"
    seasonal_mult = 1.0
    if ch == 'Paid Social' and quarter in ('Q3', 'Q4'):
        seasonal_mult = 1.4  # Diwali auction surge

    # Spend (log-normal, right-skewed)
    base_spend = np.random.lognormal(mean=np.log(50000), sigma=0.7)
    spend = max(5000, min(250000, base_spend * seasonal_mult))
    
    # Funnel metrics
    impressions = int(np.random.uniform(2000, 500000))
    clicks = max(50, int(impressions * prof['ctr'] * np.random.uniform(0.7, 1.3)))
    leads = max(10, int(clicks * np.random.uniform(0.15, 0.35)))
    conversions = max(1, int(leads * prof['conv_rate'] * np.random.uniform(0.5, 1.5)))
    
    # CAC
    cac = spend / conversions if conversions > 0 else 0
    
    # LTV & downstream
    ltv = max(4000, np.random.normal(tier_p['ltv_mean'], tier_p['ltv_std']))
    ltv_cac = ltv / cac if cac > 0 else 0
    
    monthly_rev = max(1500, np.random.normal(tier_p['rev_mean'], tier_p['rev_mean'] * 0.2))
    payback = cac / monthly_rev * 12 if monthly_rev > 0 else 24
    payback = max(3, min(24, payback))
    
    churn = max(0.03, min(0.35, np.random.normal(tier_p['churn_mean'], 0.04)))
    
    rows.append({
        'campaign_id': i + 1,
        'campaign_date': pd.Timestamp(dt).strftime('%Y-%m-%d'),
        'channel': ch,
        'city': city,
        'subscription_tier': tier,
        'spend_amount': round(spend, 2),
        'impressions': impressions,
        'clicks': clicks,
        'leads': leads,
        'conversions': conversions,
        'cac_calculated': round(cac, 2),
        'lifetime_value': round(ltv, 2),
        'ltv_cac_ratio': round(ltv_cac, 2),
        'payback_months': round(payback, 1),
        'churn_rate': round(churn, 4),
        'monthly_revenue': round(monthly_rev, 2),
        'seasonal_period': quarter,
        'campaign_status': status,
    })

df = pd.DataFrame(rows)

# ── Inject data quality issues ─────────────────────────────────────────────
# 1. ~2.5% missing values in spend_amount, cac_calculated, churn_rate
for col in ['spend_amount', 'cac_calculated', 'churn_rate']:
    mask = np.random.choice(df.index, size=int(len(df) * 0.025), replace=False)
    df.loc[mask, col] = np.nan

# 2. 8 exact duplicate rows
dup_idx = np.random.choice(df.index[:500], size=8, replace=False)
dups = df.loc[dup_idx].copy()
df = pd.concat([df, dups], ignore_index=True)

# 3. 6 extreme outliers in spend_amount (3σ+)
outlier_idx = np.random.choice(df.index, size=6, replace=False)
df.loc[outlier_idx, 'spend_amount'] = np.random.uniform(400000, 800000, size=6).round(2)

# 4. Inconsistent labels
bad_label_idx = np.random.choice(df[df['channel'] == 'Paid Search'].index, size=2, replace=False)
df.loc[bad_label_idx, 'channel'] = 'paid search'
bad_label_idx2 = np.random.choice(df[df['channel'] == 'Referral'].index, size=2, replace=False)
df.loc[bad_label_idx2, 'channel'] = 'REFERRAL'

# 5. Some campaign_date as string with different format
bad_date_idx = np.random.choice(df.index, size=5, replace=False)
for idx in bad_date_idx:
    d = df.loc[idx, 'campaign_date']
    df.loc[idx, 'campaign_date'] = pd.Timestamp(d).strftime('%d-%m-%Y')

# Shuffle rows
df = df.sample(frac=1, random_state=42).reset_index(drop=True)
# Re-assign campaign_id
df['campaign_id'] = range(1, len(df) + 1)

out_path = r'c:\Users\asus\Downloads\MERN-Expense-Tracker-main\data\cac_eda_dataset.csv'
df.to_csv(out_path, index=False)
print(f"Dataset saved: {out_path}")
print(f"Shape: {df.shape}")
print(f"Columns: {list(df.columns)}")
print(f"Null counts:\n{df.isnull().sum()[df.isnull().sum() > 0]}")
print(f"Duplicate rows: {df.duplicated().sum()}")
print(f"Unique channels: {df['channel'].unique()}")
