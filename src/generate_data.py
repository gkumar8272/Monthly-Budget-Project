import os
import sqlite3
import random
import datetime
import numpy as np
import pandas as pd

# Set deterministic random seed
SEED = 42
random.seed(SEED)
np.random.seed(SEED)

DATA_DIR = os.path.abspath("data")
DB_PATH = os.path.join(DATA_DIR, "app.db")
EXCEL_PATH = os.path.join(DATA_DIR, "cac_analysis_data.xlsx")

def generate_all_data():
    os.makedirs(DATA_DIR, exist_ok=True)
    print("=" * 70)
    print("STARTING SYNTHETIC DATASET GENERATION (STEP 10)")
    print(f"Target Database: {DB_PATH}")
    print(f"Target Excel: {EXCEL_PATH}")
    print("=" * 70)

    # -------------------------------------------------------------------------
    # 1. TABLE 1: dim_channels (Master: 5 rows)
    # -------------------------------------------------------------------------
    channels_data = [
        {"channel_id": 1, "channel_name": "Search", "channel_type": "Paid", "attribution_model": "Last-Touch", "target_cpa": 55.0, "is_active": 1},
        {"channel_id": 2, "channel_name": "Social", "channel_type": "Paid", "attribution_model": "Last-Touch", "target_cpa": 75.0, "is_active": 1},
        {"channel_id": 3, "channel_name": "Referral", "channel_type": "Partner", "attribution_model": "Last-Touch", "target_cpa": 30.0, "is_active": 1},
        {"channel_id": 4, "channel_name": "Email", "channel_type": "Organic", "attribution_model": "Last-Touch", "target_cpa": 20.0, "is_active": 1},
        {"channel_id": 5, "channel_name": "Organic", "channel_type": "Organic", "attribution_model": "Last-Touch", "target_cpa": 15.0, "is_active": 1}
    ]
    df_channels = pd.DataFrame(channels_data)

    # -------------------------------------------------------------------------
    # 2. TABLE 2: dim_subscription_tiers (Master: 4 rows)
    # -------------------------------------------------------------------------
    tiers_data = [
        {"tier_id": 1, "tier_name": "Basic", "monthly_price": 29.0, "gross_margin_pct": 0.70, "target_ltv_cac": 3.0, "annual_churn_rate": 0.24},
        {"tier_id": 2, "tier_name": "Pro", "monthly_price": 79.0, "gross_margin_pct": 0.75, "target_ltv_cac": 3.5, "annual_churn_rate": 0.18},
        {"tier_id": 3, "tier_name": "SMB", "monthly_price": 199.0, "gross_margin_pct": 0.80, "target_ltv_cac": 4.0, "annual_churn_rate": 0.12},
        {"tier_id": 4, "tier_name": "Enterprise", "monthly_price": 599.0, "gross_margin_pct": 0.85, "target_ltv_cac": 5.0, "annual_churn_rate": 0.06}
    ]
    df_tiers = pd.DataFrame(tiers_data)

    # -------------------------------------------------------------------------
    # 3. TABLE 3: fact_marketing_campaigns (240 rows: 24 months x 5 channels x 2 campaigns)
    # -------------------------------------------------------------------------
    campaigns = []
    camp_id = 1
    
    # 24 monthly periods: 2024-01 to 2025-12
    dates = pd.date_range(start="2024-01-01", periods=24, freq="MS")
    
    for dt in dates:
        ym = dt.strftime("%Y-%m")
        month_num = dt.month
        # Built-in Pattern 2: Q4 Holiday surge (Oct-Dec)
        is_q4 = month_num in [10, 11, 12]
        is_spring = month_num in [3, 4]
        seasonal_tag = "Q4_Holiday" if is_q4 else ("Spring_Promo" if is_spring else "Regular")
        surge_multiplier = 1.35 if is_q4 else (1.10 if is_spring else 1.0)
        
        for ch in channels_data:
            c_id = ch["channel_id"]
            c_name = ch["channel_name"]
            
            # Base parameters per channel
            if c_name == "Search":
                base_spend = 6500.0 * surge_multiplier
                base_impr = 120000
                ctr = 0.038
                cvr_lead = 0.14
                cvr_cust = 0.15
            elif c_name == "Social":
                base_spend = 8500.0 * surge_multiplier
                base_impr = 220000
                ctr = 0.022
                cvr_lead = 0.09
                cvr_cust = 0.04 # Pattern 4: Low conversion at checkout
            elif c_name == "Referral":
                base_spend = 1800.0
                base_impr = 25000
                ctr = 0.065
                cvr_lead = 0.28
                cvr_cust = 0.24 # High conversion
            elif c_name == "Email":
                base_spend = 1200.0
                base_impr = 45000
                ctr = 0.045
                cvr_lead = 0.18
                cvr_cust = 0.16
            else: # Organic
                base_spend = 2200.0 # Content & SEO agency overhead
                base_impr = 160000
                ctr = 0.032
                cvr_lead = 0.15
                cvr_cust = 0.18

            # Generate 2 campaign variants per channel per month (e.g. Brand vs Generic / Retargeting)
            for variant in ["Core_Acquisition", "Growth_Expansion"]:
                # Log-normal noise
                noise = np.random.lognormal(mean=0.0, sigma=0.08)
                spend = round(base_spend * 0.5 * noise, 2)
                impr = int(base_impr * 0.5 * noise)
                clicks = int(impr * ctr * np.random.uniform(0.92, 1.08))
                leads = max(int(clicks * cvr_lead * np.random.uniform(0.90, 1.10)), 1)
                conv = max(int(leads * cvr_cust * np.random.uniform(0.90, 1.10)), 1)
                
                # Intentional Outlier 1 & 2 in Social (Agency over-bid anomaly)
                if camp_id in [47, 168] and c_name == "Social":
                    spend = spend * 3.2 # Spend spike anomaly
                    
                cac = round(spend / conv, 2)
                cpc = round(spend / max(clicks, 1), 2)
                cpl = round(spend / max(leads, 1), 2)

                campaigns.append({
                    "campaign_id": camp_id,
                    "channel_id": c_id,
                    "campaign_name": f"{c_name}_{ym}_{variant}",
                    "campaign_date": dt.strftime("%Y-%m-%d"),
                    "year_month": ym,
                    "seasonal_period": seasonal_tag,
                    "spend_amount": spend,
                    "impressions": impr,
                    "clicks": clicks,
                    "leads": leads,
                    "conversions_count": conv,
                    "cac_calculated": cac,
                    "cpc_calculated": cpc,
                    "cpl_calculated": cpl
                })
                camp_id += 1

    df_campaigns = pd.DataFrame(campaigns)

    # -------------------------------------------------------------------------
    # 4. TABLE 4: fact_customer_cohorts (300 customer rows)
    # -------------------------------------------------------------------------
    # Realistic Indian Corporate / SaaS Customer Profiles
    indian_company_prefixes = [
        "Bharat", "Indus", "Aarav", "Priya", "Zenith", "Tata", "Reliance", "Kaveri", 
        "Himalaya", "Ganga", "Deccan", "Vedic", "Maha", "Surya", "Aditi", "Banyan",
        "Apex", "NxtGen", "Sankalp", "Udaan", "Swastik", "Navkar", "Shree", "Kuber"
    ]
    indian_company_suffixes = [
        "Logistics Pvt Ltd", "Retail Tech", "Infotech Solutions", "Agro Foods", 
        "HealthCare Labs", "Fintech Innovations", "Consulting Group", "Enterprises", 
        "Digital Media", "Supply Chain Ltd", "Analytics AI", "E-Commerce Hub",
        "Manufacturing Corp", "Pharma Research", "Hospitality Services", "EdTech Ventures"
    ]
    indian_cities = [
        "Bengaluru", "Mumbai", "Pune", "Gurugram", "Noida", "Hyderabad", 
        "Chennai", "Ahmedabad", "Kolkata", "Jaipur", "Chandigarh", "Indore"
    ]

    customers = []
    
    # Distribution of tiers: Basic: 35%, Pro: 35%, SMB: 20%, Enterprise: 10%
    tier_pool = [1]*105 + [2]*105 + [3]*60 + [4]*30
    random.shuffle(tier_pool)

    # Distribution of acquisition channels: Search: 30%, Social: 25%, Organic: 20%, Referral: 15%, Email: 10%
    channel_pool = [1]*90 + [2]*75 + [5]*60 + [3]*45 + [4]*30
    random.shuffle(channel_pool)

    for i in range(1, 301):
        c_code = f"CUST-IN-{i:04d}"
        ch_id = channel_pool[i - 1]
        t_id = tier_pool[i - 1]
        
        tier_row = df_tiers.loc[df_tiers["tier_id"] == t_id].iloc[0]
        mrr = float(tier_row["monthly_price"])
        margin_pct = float(tier_row["gross_margin_pct"])
        ann_churn = float(tier_row["annual_churn_rate"])

        # Acquisition date randomly between Jan 2024 and Dec 2025
        random_days = random.randint(0, 720)
        acq_date = datetime.date(2024, 1, 1) + datetime.timedelta(days=random_days)
        
        # Tenure months: based on plan tier retention profile (Enterprise retains longest)
        if t_id == 4: # Enterprise
            tenure = random.randint(12, 24)
            servicing = round(random.uniform(65.0, 95.0), 2)
            cac_base = random.uniform(280.0, 420.0)
            churn_prob = round(random.uniform(0.02, 0.08), 3)
            is_churn = 1 if random.random() < 0.05 else 0
        elif t_id == 3: # SMB
            tenure = random.randint(8, 22)
            servicing = round(random.uniform(25.0, 45.0), 2)
            cac_base = random.uniform(160.0, 260.0)
            churn_prob = round(random.uniform(0.08, 0.16), 3)
            is_churn = 1 if random.random() < 0.12 else 0
        elif t_id == 2: # Pro
            tenure = random.randint(4, 18)
            servicing = round(random.uniform(12.0, 22.0), 2)
            cac_base = random.uniform(70.0, 130.0)
            churn_prob = round(random.uniform(0.14, 0.24), 3)
            is_churn = 1 if random.random() < 0.20 else 0
        else: # Basic
            tenure = random.randint(2, 12)
            servicing = round(random.uniform(6.0, 12.0), 2)
            cac_base = random.uniform(40.0, 90.0)
            churn_prob = round(random.uniform(0.20, 0.42), 3)
            is_churn = 1 if random.random() < 0.35 else 0

        # Adjust CAC by acquisition channel (Referral/Organic cheaper, Social costlier)
        ch_mult = {1: 1.0, 2: 1.35, 3: 0.60, 4: 0.70, 5: 0.50}[ch_id]
        cac_attr = round(cac_base * ch_mult * np.random.uniform(0.92, 1.08), 2)
        
        # LTV and Payback
        monthly_gp = mrr * margin_pct
        realized_revenue = round(mrr * tenure, 2)
        ltv = round(realized_revenue * margin_pct, 2)
        ltv_cac = round(ltv / max(cac_attr, 1.0), 2)
        payback = round(cac_attr / max(monthly_gp, 1.0), 1)

        # Indian Company Name & City
        c_name = f"{random.choice(indian_company_prefixes)} {random.choice(indian_company_suffixes)}"
        city = random.choice(indian_cities)

        customers.append({
            "customer_id": i,
            "customer_code": c_code,
            "customer_name": c_name,
            "city": city,
            "channel_id": ch_id,
            "tier_id": t_id,
            "acquisition_date": acq_date.strftime("%Y-%m-%d"),
            "monthly_revenue": mrr,
            "servicing_cost": servicing,
            "cac_attributed": cac_attr,
            "tenure_months": tenure,
            "lifetime_value": ltv,
            "ltv_cac_ratio": ltv_cac,
            "payback_months": payback,
            "churn_probability": churn_prob,
            "is_churned": is_churn
        })

    # Intentional Data Quality Issues (2-3% missing, 2-3 duplicates, 1 extreme tenure outlier)
    # A. Missing values: Set servicing_cost to None for 8 rows (~2.6%)
    for null_idx in [12, 45, 88, 142, 189, 215, 260, 292]:
        customers[null_idx]["servicing_cost"] = None
        
    # B. Duplicate records: duplicate rows 15 and 90 at end as 301 and 302
    dup_1 = dict(customers[14])
    dup_1["customer_id"] = 301
    customers.append(dup_1)
    
    dup_2 = dict(customers[89])
    dup_2["customer_id"] = 302
    customers.append(dup_2)

    # C. Outlier tenure: record 55 has 48 months tenure
    customers[54]["tenure_months"] = 48
    customers[54]["lifetime_value"] = round(customers[54]["monthly_revenue"] * 48 * 0.85, 2)

    df_customers = pd.DataFrame(customers)

    # -------------------------------------------------------------------------
    # 5. CREATE SQLITE DATABASE (data/app.db) & EXECUTE DDL
    # -------------------------------------------------------------------------
    if os.path.exists(DB_PATH):
        os.remove(DB_PATH)

    conn = sqlite3.connect(DB_PATH)
    cur = conn.cursor()
    cur.execute("PRAGMA foreign_keys = ON;")

    # Table DDLs
    cur.execute("""
    CREATE TABLE dim_channels (
        channel_id INTEGER PRIMARY KEY,
        channel_name TEXT NOT NULL UNIQUE CHECK (channel_name IN ('Search', 'Social', 'Referral', 'Email', 'Organic')),
        channel_type TEXT NOT NULL CHECK (channel_type IN ('Paid', 'Organic', 'Partner')),
        attribution_model TEXT NOT NULL DEFAULT 'Last-Touch',
        target_cpa REAL NOT NULL CHECK (target_cpa > 0.0),
        is_active INTEGER NOT NULL DEFAULT 1 CHECK (is_active IN (0, 1))
    );
    """)

    cur.execute("""
    CREATE TABLE dim_subscription_tiers (
        tier_id INTEGER PRIMARY KEY,
        tier_name TEXT NOT NULL UNIQUE CHECK (tier_name IN ('Basic', 'Pro', 'SMB', 'Enterprise')),
        monthly_price REAL NOT NULL CHECK (monthly_price > 0.0),
        gross_margin_pct REAL NOT NULL CHECK (gross_margin_pct BETWEEN 0.10 AND 0.99),
        target_ltv_cac REAL NOT NULL DEFAULT 3.0 CHECK (target_ltv_cac >= 1.0),
        annual_churn_rate REAL NOT NULL CHECK (annual_churn_rate BETWEEN 0.01 AND 0.90)
    );
    """)

    cur.execute("""
    CREATE TABLE fact_marketing_campaigns (
        campaign_id INTEGER PRIMARY KEY AUTOINCREMENT,
        channel_id INTEGER NOT NULL,
        campaign_name TEXT NOT NULL,
        campaign_date TEXT NOT NULL,
        year_month TEXT NOT NULL,
        seasonal_period TEXT NOT NULL DEFAULT 'Regular' CHECK (seasonal_period IN ('Regular', 'Q4_Holiday', 'Spring_Promo')),
        spend_amount REAL NOT NULL CHECK (spend_amount >= 0.0),
        impressions INTEGER NOT NULL CHECK (impressions >= 0),
        clicks INTEGER NOT NULL CHECK (clicks >= 0),
        leads INTEGER NOT NULL CHECK (leads >= 0),
        conversions_count INTEGER NOT NULL CHECK (conversions_count >= 0),
        cac_calculated REAL NOT NULL DEFAULT 0.0,
        cpc_calculated REAL NOT NULL DEFAULT 0.0,
        cpl_calculated REAL NOT NULL DEFAULT 0.0,
        FOREIGN KEY (channel_id) REFERENCES dim_channels (channel_id) ON DELETE RESTRICT ON UPDATE CASCADE
    );
    """)

    cur.execute("""
    CREATE TABLE fact_customer_cohorts (
        customer_id INTEGER PRIMARY KEY AUTOINCREMENT,
        customer_code TEXT NOT NULL,
        customer_name TEXT,
        city TEXT,
        channel_id INTEGER NOT NULL,
        tier_id INTEGER NOT NULL,
        acquisition_date TEXT NOT NULL,
        monthly_revenue REAL NOT NULL CHECK (monthly_revenue > 0.0),
        servicing_cost REAL,
        cac_attributed REAL NOT NULL CHECK (cac_attributed >= 0.0),
        tenure_months INTEGER NOT NULL CHECK (tenure_months >= 1),
        lifetime_value REAL NOT NULL DEFAULT 0.0,
        ltv_cac_ratio REAL NOT NULL DEFAULT 0.0,
        payback_months REAL NOT NULL DEFAULT 0.0,
        churn_probability REAL NOT NULL CHECK (churn_probability BETWEEN 0.0 AND 1.0),
        is_churned INTEGER NOT NULL DEFAULT 0 CHECK (is_churned IN (0, 1)),
        FOREIGN KEY (channel_id) REFERENCES dim_channels (channel_id) ON DELETE RESTRICT ON UPDATE CASCADE,
        FOREIGN KEY (tier_id) REFERENCES dim_subscription_tiers (tier_id) ON DELETE RESTRICT ON UPDATE CASCADE
    );
    """)

    # Indices
    cur.execute("CREATE INDEX idx_camp_channel_date ON fact_marketing_campaigns (channel_id, campaign_date);")
    cur.execute("CREATE INDEX idx_cust_channel_tier ON fact_customer_cohorts (channel_id, tier_id);")

    # Insert Data into SQLite
    df_channels.to_sql("dim_channels", conn, if_exists="append", index=False)
    df_tiers.to_sql("dim_subscription_tiers", conn, if_exists="append", index=False)
    df_campaigns.to_sql("fact_marketing_campaigns", conn, if_exists="append", index=False)
    df_customers.to_sql("fact_customer_cohorts", conn, if_exists="append", index=False)

    # CREATE 34-COLUMN JOINED VIEW
    cur.execute("""
    CREATE VIEW view_cac_full_analysis AS
    SELECT 
        cu.customer_id,
        cu.customer_code,
        cu.acquisition_date,
        strftime('%Y-%m', cu.acquisition_date) AS acquisition_year_month,
        ch.channel_id,
        ch.channel_name,
        ch.channel_type,
        tr.tier_id,
        tr.tier_name,
        cu.monthly_revenue,
        tr.gross_margin_pct,
        ROUND(cu.monthly_revenue * tr.gross_margin_pct, 2) AS monthly_gross_profit,
        COALESCE(cu.servicing_cost, 0.0) AS servicing_cost,
        ROUND((cu.monthly_revenue * tr.gross_margin_pct) - COALESCE(cu.servicing_cost, 0.0), 2) AS net_contribution_margin,
        cu.cac_attributed,
        cu.tenure_months,
        cu.is_churned,
        cu.churn_probability,
        ROUND(cu.monthly_revenue * cu.tenure_months, 2) AS realized_gross_revenue,
        cu.lifetime_value,
        cu.ltv_cac_ratio,
        CASE 
            WHEN cu.ltv_cac_ratio >= 3.0 THEN 'Healthy'
            WHEN cu.ltv_cac_ratio >= 1.0 THEN 'Sub-Optimal'
            ELSE 'Value-Destroying'
        END AS ltv_health_category,
        cu.payback_months,
        CASE 
            WHEN cu.payback_months <= 12.0 THEN 'Safe (<12m)'
            ELSE 'Critical (>12m)'
        END AS payback_status,
        COALESCE(camp.spend_amount, 0.0) AS channel_monthly_spend,
        COALESCE(camp.impressions, 0) AS channel_impressions,
        COALESCE(camp.clicks, 0) AS channel_clicks,
        COALESCE(camp.leads, 0) AS channel_leads,
        COALESCE(camp.conversions_count, 0) AS channel_conversions,
        COALESCE(camp.cac_calculated, 0.0) AS channel_cac_monthly,
        CASE WHEN camp.impressions > 0 THEN ROUND((CAST(camp.clicks AS REAL) / camp.impressions) * 100.0, 2) ELSE 0.0 END AS click_through_rate,
        CASE WHEN camp.leads > 0 THEN ROUND((CAST(camp.conversions_count AS REAL) / camp.leads) * 100.0, 2) ELSE 0.0 END AS lead_conversion_rate,
        COALESCE(camp.seasonal_period, 'Regular') AS seasonal_period,
        CASE 
            WHEN CAST(strftime('%m', cu.acquisition_date) AS INTEGER) BETWEEN 1 AND 3 THEN 'Q1'
            WHEN CAST(strftime('%m', cu.acquisition_date) AS INTEGER) BETWEEN 4 AND 6 THEN 'Q2'
            WHEN CAST(strftime('%m', cu.acquisition_date) AS INTEGER) BETWEEN 7 AND 9 THEN 'Q3'
            ELSE 'Q4'
        END AS quarter_index
    FROM fact_customer_cohorts cu
    JOIN dim_channels ch ON cu.channel_id = ch.channel_id
    JOIN dim_subscription_tiers tr ON cu.tier_id = tr.tier_id
    LEFT JOIN fact_marketing_campaigns camp 
        ON cu.channel_id = camp.channel_id 
        AND strftime('%Y-%m', cu.acquisition_date) = camp.year_month;
    """)

    conn.commit()

    # Read back the view
    df_view = pd.read_sql_query("SELECT * FROM view_cac_full_analysis", conn)
    conn.close()

    # -------------------------------------------------------------------------
    # 6. EXPORT EXCEL & CSV ARTIFACTS
    # -------------------------------------------------------------------------
    # Save individual CSVs
    df_channels.to_csv(os.path.join(DATA_DIR, "dim_channels.csv"), index=False)
    df_tiers.to_csv(os.path.join(DATA_DIR, "dim_subscription_tiers.csv"), index=False)
    df_campaigns.to_csv(os.path.join(DATA_DIR, "fact_marketing_campaigns.csv"), index=False)
    df_customers.to_csv(os.path.join(DATA_DIR, "fact_customer_cohorts.csv"), index=False)
    df_view.to_csv(os.path.join(DATA_DIR, "view_cac_full_analysis.csv"), index=False)

    # Save multi-sheet Excel
    with pd.ExcelWriter(EXCEL_PATH, engine="openpyxl") as writer:
        df_channels.to_excel(writer, sheet_name="dim_channels", index=False)
        df_tiers.to_excel(writer, sheet_name="dim_subscription_tiers", index=False)
        df_campaigns.to_excel(writer, sheet_name="fact_marketing_campaigns", index=False)
        df_customers.to_excel(writer, sheet_name="fact_customer_cohorts", index=False)
        df_view.to_excel(writer, sheet_name="view_cac_full_analysis", index=False)

    print("SUCCESS: Database and Excel files generated!")
    print(f"dim_channels: {len(df_channels)} rows")
    print(f"dim_subscription_tiers: {len(df_tiers)} rows")
    print(f"fact_marketing_campaigns: {len(df_campaigns)} rows")
    print(f"fact_customer_cohorts: {len(df_customers)} rows")
    print(f"view_cac_full_analysis: {len(df_view)} rows, {df_view.shape[1]} columns")

    return df_channels, df_tiers, df_campaigns, df_customers, df_view

if __name__ == "__main__":
    generate_all_data()
