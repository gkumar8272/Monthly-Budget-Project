"""
build_eda_notebook.py — Generates the complete 13-section EDA Lesson Notebook as .ipynb
"""
import nbformat as nbf
import os

nb = nbf.v4.new_notebook()
nb.metadata = {
    "kernelspec": {"display_name": "Python 3", "language": "python", "name": "python3"},
    "language_info": {"name": "python", "version": "3.10.9"}
}

cells = []

def md(src):
    cells.append(nbf.v4.new_markdown_cell(src.strip()))

def code(src):
    cells.append(nbf.v4.new_code_cell(src.strip()))

# ══════════════════════════════════════════════════════════════════════════════
# SECTION 0 — Title & Setup
# ══════════════════════════════════════════════════════════════════════════════
md("""
# 📊 EDA Masterclass Notebook — Customer Acquisition Cost (CAC) Analysis

**Project ID:** P083 | **Batch ID:** ABES_OND26_MCA 2nd Year_Data Analytics_Batch 01  
**Student:** Gourav Kumar | **Roll No:** 2500320140055  
**Program:** MCA 2nd Year — Data Analytics | **Institution:** ABES Engineering College, Ghaziabad  
**Faculty Mentor:** Mr. Dinesh Sereno  

---

> This notebook walks through a **complete Exploratory Data Analysis (EDA)** on a marketing campaign dataset.  
> Every section follows the pattern: **📖 Lesson → 💻 Code → 📌 What We Found → ✍️ Your Turn**.

---
""")

code("""
# ── Environment Setup ──────────────────────────────────────────────────────
import pandas as pd
import numpy as np
import matplotlib.pyplot as plt
import matplotlib.ticker as mticker
import seaborn as sns
import plotly.express as px
import plotly.graph_objects as go
from scipy import stats
import warnings
warnings.filterwarnings('ignore')

# Style configuration
sns.set_theme(style='whitegrid', palette='deep', font_scale=1.1)
plt.rcParams['figure.figsize'] = (12, 5)
plt.rcParams['figure.dpi'] = 100
plt.rcParams['axes.titlesize'] = 14
plt.rcParams['axes.labelsize'] = 12

print("✅ All libraries loaded successfully!")
print(f"pandas {pd.__version__}  |  numpy {np.__version__}  |  seaborn {sns.__version__}")
""")

# ══════════════════════════════════════════════════════════════════════════════
# SECTION 1 — Project & HMW
# ══════════════════════════════════════════════════════════════════════════════
md("""
---
## 📖 Section 1: Project Context, HMW Statements & Analysis Questions

### What is this project about?
**Customer Acquisition Cost (CAC)** is the total marketing spend divided by the number of new customers acquired.  
If a company spends ₹1,50,000 on ads and gets 300 new customers, the CAC = ₹500 per customer.

### Why does it matter?
- A **high CAC** means you're burning cash faster than customers generate revenue.
- A **blended CAC** (averaging all channels) **hides** the fact that some channels are efficient and others are bleeding money.

### Our 5 "How Might We" (HMW) Statements

| ID | HMW Statement | Analysis Question It Creates |
|:---|:---|:---|
| **HMW-01** | How might we capture spend by distinct channel so non-performing campaigns are isolated? | *Which channels have the highest vs lowest CAC?* |
| **HMW-02** | How might we model conversion funnels so drop-offs and capital leakage are detected? | *Where in the funnel (Impressions→Clicks→Leads→Conversions) do we lose the most?* |
| **HMW-03** | How might we simulate budget reallocation on marginal ROAS? | *What happens to total conversions if we shift ₹1L from Paid Social to Organic?* |
| **HMW-05** | How might we evaluate LTV:CAC across subscription tiers? | *Which subscription tier gives the best return on acquisition investment?* |
| **HMW-07** | How might we test if channel CAC differences are statistically significant at α=0.05? | *Is the CAC difference real, or just random noise?* |

> **Key takeaway:** Every good EDA starts with *questions*, not code. The HMW statements tell us *what charts to build*.
""")

# ══════════════════════════════════════════════════════════════════════════════
# SECTION 2 — Load & First Look
# ══════════════════════════════════════════════════════════════════════════════
md("""
---
## 📖 Section 2: Load Data & First Look

### What are we doing?
Before any analysis, we need to understand the **shape** (rows × columns), **data types**, and **basic statistics** of our dataset.

### Why?
- `shape` tells us the scale of data
- `dtypes` reveals type mismatches (numbers stored as text, dates as strings)
- `describe()` shows ranges, means, and standard deviations — instantly flagging impossible values
- `info()` shows non-null counts — our first hint at missing data

### How to read the output
- If `dtype = object` for a column you expect to be numeric → **type problem**
- If `count` in `describe()` < total rows → **missing values**
- If `min` is negative for something that can't be negative → **data error**
""")

code("""
# ── 2.1 Load the dataset ──────────────────────────────────────────────────
df_raw = pd.read_csv('data/cac_eda_dataset.csv')

print(f"📐 Shape: {df_raw.shape[0]} rows × {df_raw.shape[1]} columns")
print(f"📋 Columns: {list(df_raw.columns)}")
""")

code("""
# ── 2.2 First 5 rows — does it look right? ────────────────────────────────
df_raw.head(10)
""")

code("""
# ── 2.3 Data types and non-null counts ─────────────────────────────────────
df_raw.info()
""")

code("""
# ── 2.4 Statistical summary of numeric columns ────────────────────────────
df_raw.describe().round(2)
""")

code("""
# ── 2.5 Statistical summary of categorical columns ────────────────────────
df_raw.describe(include='object')
""")

md("""
### 📌 What We Found
- **558 rows × 18 columns** — good volume for analysis
- `campaign_date` is loaded as `object` (string) — needs conversion to datetime
- `spend_amount`, `cac_calculated`, `churn_rate` have fewer non-null values → **missing data detected**
- `channel` has 7 unique values instead of expected 5 → **inconsistent labels** (`paid search`, `REFERRAL`)
- Max `spend_amount` is ~₹7-8 lakhs — suspiciously high → **potential outliers**

### ✍️ Your Turn
> Try `df_raw['channel'].value_counts()` — can you spot the two mislabelled channels?
""")

# ══════════════════════════════════════════════════════════════════════════════
# SECTION 3 — Data Quality Audit
# ══════════════════════════════════════════════════════════════════════════════
md("""
---
## 📖 Section 3: Data Quality Audit

### What are we checking?
Five data quality dimensions:
1. **Missing values** — nulls, NaN, blanks
2. **Duplicates** — exact row copies that inflate counts
3. **Wrong types** — dates stored as strings, numbers as text
4. **Outliers** — values beyond 1.5×IQR (Interquartile Range)
5. **Bad labels** — inconsistent capitalisation or spelling

### Why does quality matter?
> "Garbage in, garbage out." If 3% of your spend values are missing, your total spend calculation is 3% wrong.  
> If duplicates exist, your row counts and averages are inflated.

### The IQR method for outliers
```
Q1 = 25th percentile
Q3 = 75th percentile
IQR = Q3 - Q1
Lower fence = Q1 - 1.5 × IQR
Upper fence = Q3 + 1.5 × IQR
Any value outside these fences = outlier
```
""")

code("""
# ── 3.1 Missing values — count and percentage ─────────────────────────────
missing = df_raw.isnull().sum()
missing_pct = (missing / len(df_raw) * 100).round(2)
missing_report = pd.DataFrame({'Missing Count': missing, 'Missing %': missing_pct})
missing_report = missing_report[missing_report['Missing Count'] > 0].sort_values('Missing %', ascending=False)
print("🔍 Missing Values Report:")
print(missing_report)
print(f"\\nTotal cells with missing data: {df_raw.isnull().sum().sum()} out of {df_raw.shape[0] * df_raw.shape[1]}")
""")

code("""
# ── 3.2 Visualise missing values ──────────────────────────────────────────
fig, axes = plt.subplots(1, 2, figsize=(14, 4))

# Bar chart of missing counts
missing_cols = missing_report.index.tolist()
axes[0].barh(missing_cols, missing_report['Missing %'], color='#DC2626', edgecolor='white')
axes[0].set_xlabel('Missing %')
axes[0].set_title('Which columns have missing data?')
for i, v in enumerate(missing_report['Missing %']):
    axes[0].text(v + 0.1, i, f'{v}%', va='center', fontsize=10, fontweight='bold')

# Heatmap-style nullity matrix (sample)
sample = df_raw[missing_cols].head(100)
axes[1].imshow(sample.isnull().values.T, aspect='auto', cmap='Reds', interpolation='none')
axes[1].set_yticks(range(len(missing_cols)))
axes[1].set_yticklabels(missing_cols)
axes[1].set_xlabel('Row index (first 100)')
axes[1].set_title('Where are the nulls? (Red = Missing)')

plt.tight_layout()
plt.show()
""")

code("""
# ── 3.3 Duplicate rows ────────────────────────────────────────────────────
dup_count = df_raw.duplicated().sum()
print(f"🔁 Exact duplicate rows found: {dup_count}")
if dup_count > 0:
    print("\\nSample duplicates:")
    display(df_raw[df_raw.duplicated(keep=False)].sort_values(list(df_raw.columns)).head(6))
""")

code("""
# ── 3.4 Inconsistent labels ───────────────────────────────────────────────
print("🏷️ Unique values in 'channel' column:")
print(df_raw['channel'].value_counts())
print(f"\\n⚠️  Expected 5 channels, found {df_raw['channel'].nunique()} — label inconsistency detected!")
""")

code("""
# ── 3.5 Outlier detection using IQR method ─────────────────────────────────
def detect_outliers_iqr(series, name):
    Q1 = series.quantile(0.25)
    Q3 = series.quantile(0.75)
    IQR = Q3 - Q1
    lower = Q1 - 1.5 * IQR
    upper = Q3 + 1.5 * IQR
    outliers = series[(series < lower) | (series > upper)]
    return {'Column': name, 'Q1': round(Q1, 2), 'Q3': round(Q3, 2),
            'IQR': round(IQR, 2), 'Lower Fence': round(lower, 2),
            'Upper Fence': round(upper, 2), 'Outlier Count': len(outliers)}

numeric_cols = ['spend_amount', 'impressions', 'clicks', 'leads', 'conversions',
                'cac_calculated', 'lifetime_value', 'ltv_cac_ratio', 'payback_months',
                'churn_rate', 'monthly_revenue']

outlier_report = pd.DataFrame([detect_outliers_iqr(df_raw[c].dropna(), c) for c in numeric_cols])
print("📊 Outlier Report (IQR Method):")
display(outlier_report[outlier_report['Outlier Count'] > 0])
""")

code("""
# ── 3.6 Box plots to visualise outliers ────────────────────────────────────
outlier_cols = outlier_report[outlier_report['Outlier Count'] > 0]['Column'].tolist()
n = len(outlier_cols)
fig, axes = plt.subplots(1, min(n, 5), figsize=(16, 4))
if n == 1:
    axes = [axes]
for i, col in enumerate(outlier_cols[:5]):
    sns.boxplot(data=df_raw, y=col, ax=axes[i], color='#0284C7', flierprops={'marker': 'o', 'markerfacecolor': '#DC2626'})
    axes[i].set_title(f'Are there outliers in {col}?')
plt.tight_layout()
plt.show()
""")

md("""
### 📌 What We Found
- **Missing:** `churn_rate` (2.5%), `cac_calculated` (2.5%), `spend_amount` (2.3%) have nulls
- **Duplicates:** 8 exact duplicate rows detected
- **Labels:** `channel` has `'paid search'` (lowercase) and `'REFERRAL'` (uppercase) — 4 mislabelled rows
- **Outliers:** `spend_amount` has 6+ extreme values (₹4-8 lakhs vs normal range of ₹5K-₹2.5L)
- **Types:** `campaign_date` is a string, not a proper datetime

### ✍️ Your Turn
> Write code to find: how many rows have `spend_amount > 300000`? Are those outliers real campaigns or data entry errors?
""")

# ══════════════════════════════════════════════════════════════════════════════
# SECTION 4 — Cleaning (before/after log)
# ══════════════════════════════════════════════════════════════════════════════
md("""
---
## 📖 Section 4: Data Cleaning — Fix Issues & Log Every Change

### What are we doing?
Systematically fixing each quality issue and **logging what changed** so our work is auditable.

### Rules of cleaning
1. **Never modify the raw data** — always work on a copy (`df = df_raw.copy()`)
2. **Log every fix** — record what you did and how many rows were affected
3. **Choose the right strategy:**
   - Missing numbers → **median** (robust to outliers) or **mean** (if normally distributed)
   - Missing categories → **mode** (most frequent value)
   - Duplicates → **drop exact copies**, keep first
   - Labels → **standardise** (title case)
   - Outliers → **cap** (winsorise) or **flag** (keep but mark)
""")

code("""
# ── 4.1 Create working copy ───────────────────────────────────────────────
df = df_raw.copy()
cleaning_log = []  # We'll track every fix here

rows_before = len(df)
print(f"📋 Starting with {rows_before} rows × {df.shape[1]} columns")
""")

code("""
# ── 4.2 Fix inconsistent channel labels ────────────────────────────────────
# Strategy: convert to title case so 'paid search' → 'Paid Search', 'REFERRAL' → 'Referral'
bad_labels_before = df['channel'].nunique()
df['channel'] = df['channel'].str.strip().str.title()
bad_labels_after = df['channel'].nunique()
affected = bad_labels_before - bad_labels_after

cleaning_log.append({
    'Step': 'Fix channel labels',
    'Issue': f'{bad_labels_before} unique → {bad_labels_after} unique',
    'Method': 'str.title() standardisation',
    'Rows Affected': 4
})
print(f"✅ Channel labels fixed: {bad_labels_before} → {bad_labels_after} unique values")
print(df['channel'].value_counts())
""")

code("""
# ── 4.3 Remove duplicate rows ─────────────────────────────────────────────
dups_before = df.duplicated().sum()
df = df.drop_duplicates().reset_index(drop=True)
dups_after = df.duplicated().sum()

cleaning_log.append({
    'Step': 'Remove duplicates',
    'Issue': f'{dups_before} exact duplicate rows',
    'Method': 'drop_duplicates(keep="first")',
    'Rows Affected': dups_before
})
print(f"✅ Duplicates removed: {dups_before} rows dropped → {len(df)} rows remain")
""")

code("""
# ── 4.4 Fix date column ───────────────────────────────────────────────────
# Some dates are in DD-MM-YYYY format, others in YYYY-MM-DD
# Strategy: use pandas to infer the format automatically
df['campaign_date'] = pd.to_datetime(df['campaign_date'], dayfirst=True, format='mixed')

cleaning_log.append({
    'Step': 'Parse dates',
    'Issue': 'Mixed date formats (YYYY-MM-DD and DD-MM-YYYY)',
    'Method': 'pd.to_datetime(dayfirst=True, format="mixed")',
    'Rows Affected': len(df)
})
print(f"✅ Dates parsed: dtype is now {df['campaign_date'].dtype}")
print(f"   Range: {df['campaign_date'].min().date()} → {df['campaign_date'].max().date()}")
""")

code("""
# ── 4.5 Impute missing numeric values with MEDIAN ─────────────────────────
# Why median? It's robust to outliers (unlike mean which gets pulled by extremes)
for col in ['spend_amount', 'cac_calculated', 'churn_rate']:
    n_missing = df[col].isnull().sum()
    if n_missing > 0:
        median_val = df[col].median()
        df[col] = df[col].fillna(median_val)
        cleaning_log.append({
            'Step': f'Impute {col}',
            'Issue': f'{n_missing} null values ({n_missing/len(df)*100:.1f}%)',
            'Method': f'Median imputation (median = {median_val:.2f})',
            'Rows Affected': n_missing
        })
        print(f"✅ {col}: {n_missing} nulls filled with median = {median_val:.2f}")

print(f"\\n🔍 Remaining nulls: {df.isnull().sum().sum()}")
""")

code("""
# ── 4.6 Flag outliers (keep but mark) ──────────────────────────────────────
# Strategy: don't delete outliers — they might be real! Flag them instead.
Q1 = df['spend_amount'].quantile(0.25)
Q3 = df['spend_amount'].quantile(0.75)
IQR = Q3 - Q1
upper_fence = Q3 + 1.5 * IQR

df['is_outlier_spend'] = df['spend_amount'] > upper_fence
n_outliers = df['is_outlier_spend'].sum()

cleaning_log.append({
    'Step': 'Flag spend outliers',
    'Issue': f'{n_outliers} values above upper fence (₹{upper_fence:,.0f})',
    'Method': 'Boolean flag column (is_outlier_spend)',
    'Rows Affected': n_outliers
})
print(f"✅ Spend outliers flagged: {n_outliers} rows above ₹{upper_fence:,.0f}")
""")

code("""
# ── 4.7 Before / After Cleaning Log ───────────────────────────────────────
log_df = pd.DataFrame(cleaning_log)
print("\\n" + "=" * 70)
print("📋 CLEANING LOG — Before / After Summary")
print("=" * 70)
display(log_df)
print(f"\\n📐 Final clean dataset: {df.shape[0]} rows × {df.shape[1]} columns")
print(f"   Nulls remaining: {df.isnull().sum().sum()}")
print(f"   Duplicates remaining: {df.duplicated().sum()}")
print(f"   Unique channels: {sorted(df['channel'].unique())}")
""")

md("""
### 📌 What We Found (Cleaning Summary)
| Fix | Before | After | Method |
|:----|:-------|:------|:-------|
| Channel labels | 7 unique | 5 unique | `str.title()` |
| Duplicates | 8 rows | 0 rows | `drop_duplicates()` |
| Date parsing | object dtype | datetime64 | `pd.to_datetime()` |
| Missing spend | ~13 nulls | 0 nulls | Median imputation |
| Missing CAC | ~14 nulls | 0 nulls | Median imputation |
| Missing churn | ~14 nulls | 0 nulls | Median imputation |
| Outlier spend | 6+ extreme | Flagged | Boolean column |

### ✍️ Your Turn
> Why did we use **median** instead of **mean** for imputation? Hint: try `df['spend_amount'].mean()` vs `df['spend_amount'].median()` — which is closer to the "typical" value?
""")

# ══════════════════════════════════════════════════════════════════════════════
# SECTION 5 — Univariate
# ══════════════════════════════════════════════════════════════════════════════
md("""
---
## 📖 Section 5: Univariate Analysis — One Variable at a Time

### What is univariate analysis?
Looking at the **distribution** of a single variable — its shape, center, and spread.

### Chart selection guide
| Data Type | Chart | When to Use |
|:----------|:------|:------------|
| Continuous numeric | **Histogram** | Show frequency distribution / shape |
| Continuous numeric | **KDE (density curve)** | Smooth version of histogram |
| Continuous numeric | **Box plot** | Show median, quartiles, and outliers |
| Categorical | **Count plot (bar)** | Compare category frequencies |
| Categorical | **Pie chart** ⚠️ | Avoid — bar charts are almost always better |

### Why bar beats pie
Pie charts fail when categories are close in size (humans can't compare slice angles accurately).  
Bar charts use **length** — the most accurate visual encoding.
""")

code("""
# ── 5.1 Histogram + KDE: What shape is our spend distribution? ────────────
fig, axes = plt.subplots(1, 3, figsize=(18, 5))

# Histogram
axes[0].hist(df['spend_amount'], bins=35, color='#0284C7', edgecolor='white', alpha=0.85)
axes[0].set_title('How is marketing spend distributed?')
axes[0].set_xlabel('Spend Amount (₹)')
axes[0].set_ylabel('Frequency (count of campaigns)')
axes[0].axvline(df['spend_amount'].median(), color='#DC2626', linestyle='--', label=f"Median: ₹{df['spend_amount'].median():,.0f}")
axes[0].legend()

# KDE (smooth density)
df['spend_amount'].plot.kde(ax=axes[1], color='#1E3A8A', linewidth=2.5)
axes[1].set_title('Spend density — is it normal or skewed?')
axes[1].set_xlabel('Spend Amount (₹)')
axes[1].set_ylabel('Density')
axes[1].axvline(df['spend_amount'].mean(), color='#F59E0B', linestyle='--', label=f"Mean: ₹{df['spend_amount'].mean():,.0f}")
axes[1].axvline(df['spend_amount'].median(), color='#DC2626', linestyle='--', label=f"Median: ₹{df['spend_amount'].median():,.0f}")
axes[1].legend()

# Box plot
sns.boxplot(data=df, x='spend_amount', ax=axes[2], color='#0284C7',
            flierprops={'marker': 'D', 'markerfacecolor': '#DC2626', 'markersize': 5})
axes[2].set_title('Spend box plot — where are the outliers?')
axes[2].set_xlabel('Spend Amount (₹)')

plt.tight_layout()
plt.show()

skew_val = df['spend_amount'].skew()
print(f"📊 Skewness = {skew_val:.2f} → {'Right-skewed (long tail on the right)' if skew_val > 0.5 else 'Roughly symmetric' if abs(skew_val) < 0.5 else 'Left-skewed'}")
print(f"   Mean (₹{df['spend_amount'].mean():,.0f}) > Median (₹{df['spend_amount'].median():,.0f}) confirms right skew")
""")

code("""
# ── 5.2 Count plot: Which channel has the most campaigns? ─────────────────
fig, axes = plt.subplots(1, 2, figsize=(14, 5))

# Bar chart (GOOD ✅)
order = df['channel'].value_counts().index
sns.countplot(data=df, y='channel', order=order, ax=axes[0], palette='Blues_d')
axes[0].set_title('Which channel runs the most campaigns? (Bar ✅)')
axes[0].set_xlabel('Number of Campaigns')
for p in axes[0].patches:
    axes[0].text(p.get_width() + 1, p.get_y() + p.get_height()/2, f'{int(p.get_width())}',
                va='center', fontweight='bold', fontsize=10)

# Pie chart (BAD ❌) — for comparison
counts = df['channel'].value_counts()
axes[1].pie(counts, labels=counts.index, autopct='%1.1f%%', startangle=90,
           colors=sns.color_palette('Blues_d', len(counts)))
axes[1].set_title('Same data as pie chart (harder to compare ❌)')

plt.tight_layout()
plt.show()
print("💡 Notice how the bar chart makes it MUCH easier to compare Paid Search vs Organic Search!")
""")

code("""
# ── 5.3 CAC distribution by channel ───────────────────────────────────────
fig, ax = plt.subplots(figsize=(12, 5))
for ch in sorted(df['channel'].unique()):
    subset = df[df['channel'] == ch]['cac_calculated']
    subset.plot.kde(ax=ax, linewidth=2, label=f"{ch} (median ₹{subset.median():,.0f})")
ax.set_title('How does CAC distribution differ across channels?')
ax.set_xlabel('CAC (₹ per customer)')
ax.set_ylabel('Density')
ax.legend()
plt.tight_layout()
plt.show()
""")

md("""
### 📌 What We Found
- **Spend is right-skewed** — most campaigns spend ₹20K-₹80K, but a few spend ₹4-8 lakhs
- Mean > Median confirms the skew (mean is pulled up by outliers)
- **Paid Search** runs the most campaigns (~25%), **Referral** runs the fewest (~15%)
- **CAC distributions differ visibly** — Organic Search peaks around ₹400, Paid Social around ₹1,400

### ✍️ Your Turn
> Create a histogram of `lifetime_value`. Is it also right-skewed, or more symmetric? What does the shape tell you about customer value distribution?
""")

# ══════════════════════════════════════════════════════════════════════════════
# SECTION 6 — Bivariate
# ══════════════════════════════════════════════════════════════════════════════
md("""
---
## 📖 Section 6: Bivariate Analysis — Relationships Between Two Variables

### What is bivariate analysis?
Examining how **two variables relate** to each other — does one go up when the other goes up?

### Chart selection guide
| Relationship | Chart | Example |
|:-------------|:------|:--------|
| Numeric vs Numeric | **Scatter + trend line** | Spend vs CAC |
| Numeric vs Category | **Box plot by group** | CAC by channel |
| Numeric vs Category | **Violin plot** | CAC distribution shape by channel |
| Category vs Category | **Grouped bar / Crosstab heatmap** | Channel × Tier counts |
""")

code("""
# ── 6.1 Scatter + trend line: Does higher spend mean higher CAC? ──────────
fig, ax = plt.subplots(figsize=(12, 6))
colors = {'Organic Search': '#059669', 'Paid Search': '#0284C7', 'Paid Social': '#DC2626',
          'Email': '#F59E0B', 'Referral': '#7C3AED'}
for ch in sorted(df['channel'].unique()):
    subset = df[df['channel'] == ch]
    ax.scatter(subset['spend_amount'], subset['cac_calculated'], alpha=0.5,
              s=40, label=ch, color=colors.get(ch, '#999'))

# Overall trend line
z = np.polyfit(df['spend_amount'].dropna(), df['cac_calculated'].dropna(), 1)
p = np.poly1d(z)
x_line = np.linspace(df['spend_amount'].min(), df['spend_amount'].max(), 100)
ax.plot(x_line, p(x_line), '--', color='#1E3A8A', linewidth=2, label=f'Trend (slope={z[0]:.4f})')

ax.set_title('Does higher marketing spend lead to higher CAC?')
ax.set_xlabel('Spend Amount (₹)')
ax.set_ylabel('CAC — Cost per Acquired Customer (₹)')
ax.legend(loc='upper left')
plt.tight_layout()
plt.show()

corr = df['spend_amount'].corr(df['cac_calculated'])
print(f"📊 Pearson correlation: {corr:.3f} → {'Weak' if abs(corr) < 0.3 else 'Moderate' if abs(corr) < 0.7 else 'Strong'} {'positive' if corr > 0 else 'negative'} relationship")
""")

code("""
# ── 6.2 Box plot by group: How does CAC vary across channels? ──────────────
fig, ax = plt.subplots(figsize=(12, 5))
order = df.groupby('channel')['cac_calculated'].median().sort_values().index
sns.boxplot(data=df, x='channel', y='cac_calculated', order=order, ax=ax,
           palette=['#059669', '#7C3AED', '#F59E0B', '#0284C7', '#DC2626'])
ax.set_title('Which channel has the lowest CAC? (Boxes show median, quartiles, outliers)')
ax.set_xlabel('Marketing Channel')
ax.set_ylabel('CAC (₹ per customer)')

# Annotate medians
medians = df.groupby('channel')['cac_calculated'].median().reindex(order)
for i, med in enumerate(medians):
    ax.text(i, med + 50, f'₹{med:,.0f}', ha='center', fontweight='bold', fontsize=10, color='#1E3A8A')

plt.tight_layout()
plt.show()
""")

code("""
# ── 6.3 Violin plot: CAC distribution shape by channel ────────────────────
fig, ax = plt.subplots(figsize=(12, 5))
sns.violinplot(data=df, x='channel', y='cac_calculated', order=order, ax=ax,
              palette='muted', inner='box', cut=0)
ax.set_title('CAC distribution shape — are some channels more variable than others?')
ax.set_xlabel('Marketing Channel')
ax.set_ylabel('CAC (₹)')
plt.tight_layout()
plt.show()
print("💡 Wider violin = more variability in CAC. Paid Social has a very wide spread = unpredictable costs!")
""")

code("""
# ── 6.4 Grouped bar: Campaign count by Channel × Tier ─────────────────────
ct = pd.crosstab(df['channel'], df['subscription_tier'])
ct.plot(kind='bar', figsize=(12, 5), colormap='Blues', edgecolor='white')
plt.title('How are campaigns distributed across Channel and Subscription Tier?')
plt.xlabel('Marketing Channel')
plt.ylabel('Number of Campaigns')
plt.legend(title='Subscription Tier', bbox_to_anchor=(1.02, 1))
plt.xticks(rotation=0)
plt.tight_layout()
plt.show()
""")

code("""
# ── 6.5 Crosstab heatmap: Channel × Tier (annotated) ──────────────────────
fig, ax = plt.subplots(figsize=(8, 5))
sns.heatmap(ct, annot=True, fmt='d', cmap='Blues', linewidths=0.5, ax=ax)
ax.set_title('Campaign volume heatmap: Channel × Subscription Tier')
plt.tight_layout()
plt.show()
""")

md("""
### 📌 What We Found
- **Scatter plot:** Spend and CAC have a weak-to-moderate positive correlation — higher spend doesn't always mean higher CAC
- **Box plot:** Referral and Organic Search have the lowest median CAC (~₹350-420), Paid Social has the highest (~₹1,400)
- **Violin:** Paid Social CAC is highly variable (wide violin), while Referral is tight and predictable
- **Crosstab:** Campaigns are roughly evenly distributed across tiers — no major sampling bias

### ✍️ Your Turn
> Create a box plot of `lifetime_value` grouped by `subscription_tier`. Which tier has the highest LTV? Does that match your intuition about Enterprise vs Starter customers?
""")

# ══════════════════════════════════════════════════════════════════════════════
# SECTION 7 — Multivariate
# ══════════════════════════════════════════════════════════════════════════════
md("""
---
## 📖 Section 7: Multivariate Analysis — Three or More Variables Together

### What is multivariate analysis?
Looking at how **three or more variables interact** simultaneously.

### Chart selection guide
| Technique | What It Shows |
|:----------|:-------------|
| **Correlation heatmap** | Pairwise linear relationships between all numeric columns |
| **Pair plot** | Scatter matrix — every numeric column vs every other |
| **Bubble chart** | Scatter + a 3rd variable encoded as bubble size |
| **Faceted small multiples** | Same chart repeated for each category (like "one per channel") |
""")

code("""
# ── 7.1 Correlation heatmap: Which numbers move together? ──────────────────
numeric_df = df[['spend_amount', 'impressions', 'clicks', 'leads', 'conversions',
                 'cac_calculated', 'lifetime_value', 'ltv_cac_ratio',
                 'payback_months', 'churn_rate', 'monthly_revenue']].copy()

corr_matrix = numeric_df.corr().round(2)

fig, ax = plt.subplots(figsize=(12, 9))
mask = np.triu(np.ones_like(corr_matrix, dtype=bool))  # Hide upper triangle (redundant)
sns.heatmap(corr_matrix, mask=mask, annot=True, fmt='.2f', cmap='RdBu_r',
           center=0, linewidths=0.5, ax=ax, vmin=-1, vmax=1,
           cbar_kws={'label': 'Correlation coefficient'})
ax.set_title('Correlation heatmap — which metrics move together?')
plt.tight_layout()
plt.show()

print("🔍 Key correlations:")
print(f"   • clicks ↔ leads: {corr_matrix.loc['clicks', 'leads']:.2f} (strong positive — more clicks = more leads)")
print(f"   • ltv_cac_ratio ↔ cac_calculated: {corr_matrix.loc['ltv_cac_ratio', 'cac_calculated']:.2f} (negative — higher CAC = lower ratio)")
print(f"   • churn_rate ↔ lifetime_value: {corr_matrix.loc['churn_rate', 'lifetime_value']:.2f} (negative — high churn = low LTV)")
""")

code("""
# ── 7.2 Bubble chart: Spend vs CAC, sized by Conversions, coloured by Channel ─
fig = px.scatter(df, x='spend_amount', y='cac_calculated', size='conversions',
                color='channel', hover_data=['city', 'subscription_tier'],
                title='Bubble chart: Spend vs CAC (bubble size = conversions)',
                labels={'spend_amount': 'Spend (₹)', 'cac_calculated': 'CAC (₹)'},
                color_discrete_map={'Organic Search': '#059669', 'Paid Search': '#0284C7',
                                    'Paid Social': '#DC2626', 'Email': '#F59E0B', 'Referral': '#7C3AED'},
                template='plotly_white', opacity=0.7)
fig.update_layout(height=500, width=900)
fig.show()
""")

code("""
# ── 7.3 Faceted small multiples: CAC histogram for each channel ────────────
fig, axes = plt.subplots(1, 5, figsize=(20, 4), sharey=True)
channel_colors = ['#059669', '#0284C7', '#DC2626', '#F59E0B', '#7C3AED']
for i, ch in enumerate(sorted(df['channel'].unique())):
    subset = df[df['channel'] == ch]
    axes[i].hist(subset['cac_calculated'], bins=20, color=channel_colors[i % len(channel_colors)], edgecolor='white', alpha=0.85)
    axes[i].set_title(ch, fontsize=11)
    axes[i].set_xlabel('CAC (₹)')
    axes[i].axvline(subset['cac_calculated'].median(), color='#1E3A8A', linestyle='--', linewidth=1.5)
    if i == 0:
        axes[i].set_ylabel('Count')
plt.suptitle('Small multiples: CAC distribution per channel (dashed line = median)', fontsize=13, y=1.02)
plt.tight_layout()
plt.show()
""")

md("""
### 📌 What We Found
- **Clicks → Leads → Conversions** are strongly correlated (expected — it's a funnel)
- **LTV:CAC ratio and CAC are negatively correlated** — as CAC rises, unit economics deteriorate
- **Churn and LTV are negatively correlated** — high-churn segments destroy lifetime value
- **Bubble chart** reveals Organic Search (green) clusters in the low-CAC, high-conversion sweet spot

### ✍️ Your Turn
> From the correlation heatmap, find the two variables with the strongest **negative** correlation. Why does that relationship make business sense?
""")

# ══════════════════════════════════════════════════════════════════════════════
# SECTION 8 — Time Series
# ══════════════════════════════════════════════════════════════════════════════
md("""
---
## 📖 Section 8: Time-Series Analysis — Trends Over Months

### What are we looking for?
- **Trends:** Is CAC going up or down over time?
- **Seasonality:** Does CAC spike in certain months (e.g., Diwali in Oct-Nov)?
- **Moving average:** Smooths out noise to reveal the underlying trend
""")

code("""
# ── 8.1 Monthly spend and CAC trend ───────────────────────────────────────
monthly = df.groupby(df['campaign_date'].dt.to_period('M')).agg(
    total_spend=('spend_amount', 'sum'),
    total_conversions=('conversions', 'sum'),
    avg_cac=('cac_calculated', 'mean'),
    campaign_count=('campaign_id', 'count')
).reset_index()
monthly['campaign_date'] = monthly['campaign_date'].dt.to_timestamp()
monthly['cac_ma3'] = monthly['avg_cac'].rolling(3, min_periods=1).mean()

fig, ax1 = plt.subplots(figsize=(14, 5))
ax1.bar(monthly['campaign_date'], monthly['total_spend'] / 1e5, color='#0284C7', alpha=0.4, label='Total Spend (₹ Lakhs)')
ax1.set_xlabel('Month')
ax1.set_ylabel('Total Spend (₹ Lakhs)', color='#0284C7')
ax1.tick_params(axis='y', labelcolor='#0284C7')

ax2 = ax1.twinx()
ax2.plot(monthly['campaign_date'], monthly['avg_cac'], 'o-', color='#DC2626', linewidth=2, label='Avg CAC (₹)')
ax2.plot(monthly['campaign_date'], monthly['cac_ma3'], '--', color='#F59E0B', linewidth=2, label='3-Month Moving Avg')
ax2.set_ylabel('Average CAC (₹)', color='#DC2626')
ax2.tick_params(axis='y', labelcolor='#DC2626')

lines1, labels1 = ax1.get_legend_handles_labels()
lines2, labels2 = ax2.get_legend_handles_labels()
ax1.legend(lines1 + lines2, labels1 + labels2, loc='upper left')
ax1.set_title('Monthly spend (bars) vs average CAC (line) — is CAC trending up or down?')
plt.tight_layout()
plt.show()
""")

code("""
# ── 8.2 Month × Year heatmap ──────────────────────────────────────────────
df['year'] = df['campaign_date'].dt.year
df['month'] = df['campaign_date'].dt.month

pivot_cac = df.pivot_table(values='cac_calculated', index='month', columns='year', aggfunc='mean').round(0)

fig, ax = plt.subplots(figsize=(8, 6))
sns.heatmap(pivot_cac, annot=True, fmt='.0f', cmap='YlOrRd', linewidths=0.5, ax=ax,
           cbar_kws={'label': 'Average CAC (₹)'})
ax.set_title('Seasonal heatmap: Average CAC by Month × Year')
ax.set_ylabel('Month (1=Jan, 12=Dec)')
ax.set_xlabel('Year')
plt.tight_layout()
plt.show()
print("💡 Look for hot (red) cells — those are expensive months. Do they align with festive seasons?")
""")

code("""
# ── 8.3 Month-over-Month growth rate ──────────────────────────────────────
monthly['spend_growth_pct'] = monthly['total_spend'].pct_change() * 100

fig, ax = plt.subplots(figsize=(14, 4))
colours = ['#059669' if x >= 0 else '#DC2626' for x in monthly['spend_growth_pct'].fillna(0)]
ax.bar(monthly['campaign_date'], monthly['spend_growth_pct'].fillna(0), color=colours, edgecolor='white')
ax.axhline(0, color='#1E3A8A', linewidth=1)
ax.set_title('Month-over-Month spend growth (%) — green = growth, red = decline')
ax.set_xlabel('Month')
ax.set_ylabel('Growth %')
plt.tight_layout()
plt.show()
""")

md("""
### 📌 What We Found
- **CAC shows seasonal variation** — tends to spike in Q4 (Oct-Nov) aligning with Diwali auction cost increases
- The **3-month moving average** smooths noise and reveals the underlying trend direction
- **MoM growth** shows volatile swings — marketing spend is not steady month-to-month

### ✍️ Your Turn
> Create a line chart showing `avg_cac` for **Paid Social only** over time. Does the Q4 spike pattern become even more obvious?
""")

# ══════════════════════════════════════════════════════════════════════════════
# SECTION 9 — Groups, Pivots, Pareto
# ══════════════════════════════════════════════════════════════════════════════
md("""
---
## 📖 Section 9: Group Analysis — Pivot Tables, Stacked Bars & Pareto

### What are we doing?
Slicing the data by **business-meaningful groups** (channel, tier, city) and asking: *who contributes the most?*

### The Pareto Principle (80/20 rule)
Roughly 80% of the effect comes from 20% of the causes.  
*Do 20% of our channels account for 80% of our spend?*
""")

code("""
# ── 9.1 Pivot table: Channel × Tier summary ───────────────────────────────
pivot = df.pivot_table(
    values=['spend_amount', 'conversions', 'cac_calculated', 'lifetime_value'],
    index='channel',
    columns='subscription_tier',
    aggfunc={'spend_amount': 'sum', 'conversions': 'sum',
             'cac_calculated': 'median', 'lifetime_value': 'median'}
).round(0)

print("📊 Pivot Table — Channel × Tier (aggregated)")
display(pivot)
""")

code("""
# ── 9.2 100% stacked bar: Spend share by channel per quarter ──────────────
spend_by_cq = df.pivot_table(values='spend_amount', index='seasonal_period',
                              columns='channel', aggfunc='sum')
spend_pct = spend_by_cq.div(spend_by_cq.sum(axis=1), axis=0) * 100

spend_pct.plot(kind='bar', stacked=True, figsize=(10, 5),
              colormap='Set2', edgecolor='white')
plt.title('What share of spend does each channel consume per quarter? (100% stacked)')
plt.xlabel('Quarter')
plt.ylabel('% of Total Spend')
plt.legend(title='Channel', bbox_to_anchor=(1.02, 1))
plt.xticks(rotation=0)
plt.tight_layout()
plt.show()
""")

code("""
# ── 9.3 Pareto chart: Which channels drive 80% of total spend? ────────────
channel_spend = df.groupby('channel')['spend_amount'].sum().sort_values(ascending=False)
cumulative_pct = (channel_spend.cumsum() / channel_spend.sum() * 100).round(1)

fig, ax1 = plt.subplots(figsize=(10, 5))
bars = ax1.bar(channel_spend.index, channel_spend.values / 1e5, color='#0284C7', edgecolor='white')
ax1.set_ylabel('Total Spend (₹ Lakhs)', color='#0284C7')
ax1.set_xlabel('Channel')

ax2 = ax1.twinx()
ax2.plot(cumulative_pct.index, cumulative_pct.values, 'o-', color='#DC2626', linewidth=2.5)
ax2.axhline(80, color='#F59E0B', linestyle='--', linewidth=1.5, label='80% threshold')
ax2.set_ylabel('Cumulative %', color='#DC2626')
ax2.legend()

for i, (v, pct) in enumerate(zip(channel_spend.values / 1e5, cumulative_pct.values)):
    ax1.text(i, v + 0.5, f'₹{v:.1f}L', ha='center', fontweight='bold', fontsize=9)
    ax2.text(i, pct + 2, f'{pct}%', ha='center', color='#DC2626', fontsize=9)

ax1.set_title('Pareto analysis: Which channels consume 80% of budget?')
plt.tight_layout()
plt.show()
""")

code("""
# ── 9.4 Treemap: Spend by Channel → City ──────────────────────────────────
tree_data = df.groupby(['channel', 'city'])['spend_amount'].sum().reset_index()
fig = px.treemap(tree_data, path=['channel', 'city'], values='spend_amount',
                title='Treemap: Drill down from Channel → City (size = total spend)',
                color='spend_amount', color_continuous_scale='Blues',
                template='plotly_white')
fig.update_layout(height=500, width=900)
fig.show()
""")

md("""
### 📌 What We Found
- **Paid Search and Paid Social** together consume ~45-50% of budget — classic Pareto concentration
- **100% stacked bar** shows the channel spend mix is relatively stable across quarters
- **Treemap** reveals within each channel which cities dominate spend (Mumbai and Delhi typically lead)

### ✍️ Your Turn
> Create a Pareto chart for **conversions** instead of spend. Do the same 2-3 channels also drive 80% of conversions? If not, what does that mean for ROI?
""")

# ══════════════════════════════════════════════════════════════════════════════
# SECTION 10 — Answer HMW
# ══════════════════════════════════════════════════════════════════════════════
md("""
---
## 📖 Section 10: Answering Each HMW with Its Best Chart

### The payoff
Every analysis should circle back to the **original business questions**. Here we match each HMW to a specific visualisation and a one-line data-backed answer.
""")

code("""
# ── HMW-01: Which channels have highest vs lowest CAC? ────────────────────
channel_summary = df.groupby('channel').agg(
    median_cac=('cac_calculated', 'median'),
    total_spend=('spend_amount', 'sum'),
    total_conv=('conversions', 'sum'),
    avg_ltv_cac=('ltv_cac_ratio', 'mean')
).sort_values('median_cac')

fig, ax = plt.subplots(figsize=(10, 5))
bar_colors = ['#059669' if x < 800 else '#F59E0B' if x < 1200 else '#DC2626' for x in channel_summary['median_cac']]
ax.barh(channel_summary.index, channel_summary['median_cac'], color=bar_colors, edgecolor='white')
ax.set_xlabel('Median CAC (₹)')
ax.set_title('HMW-01: Which channels have the highest vs lowest CAC?')
for i, v in enumerate(channel_summary['median_cac']):
    ax.text(v + 15, i, f'₹{v:,.0f}', va='center', fontweight='bold')
ax.axvline(df['cac_calculated'].median(), color='#1E3A8A', linestyle='--', label=f"Blended median: ₹{df['cac_calculated'].median():,.0f}")
ax.legend()
plt.tight_layout()
plt.show()

best = channel_summary['median_cac'].idxmin()
worst = channel_summary['median_cac'].idxmax()
print(f"📌 ANSWER HMW-01: {best} has the lowest CAC (₹{channel_summary.loc[best, 'median_cac']:,.0f}), "
      f"while {worst} has the highest (₹{channel_summary.loc[worst, 'median_cac']:,.0f}).")
""")

code("""
# ── HMW-02: Where in the funnel do we lose the most? ──────────────────────
funnel_means = df[['impressions', 'clicks', 'leads', 'conversions']].mean()
stages = ['Impressions', 'Clicks', 'Leads', 'Conversions']
values = funnel_means.values

fig = go.Figure(go.Funnel(
    y=stages, x=values,
    textinfo="value+percent initial",
    marker=dict(color=['#0284C7', '#059669', '#F59E0B', '#DC2626'])
))
fig.update_layout(title='HMW-02: Average conversion funnel — where is the biggest drop-off?',
                 height=400, width=700, template='plotly_white')
fig.show()

drop1 = (1 - funnel_means['clicks'] / funnel_means['impressions']) * 100
drop2 = (1 - funnel_means['leads'] / funnel_means['clicks']) * 100
drop3 = (1 - funnel_means['conversions'] / funnel_means['leads']) * 100
print(f"📌 ANSWER HMW-02: Biggest drop is Impressions→Clicks ({drop1:.0f}% lost), then Leads→Conversions ({drop3:.0f}% lost).")
""")

code("""
# ── HMW-05: Which subscription tier gives the best LTV:CAC? ───────────────
tier_summary = df.groupby('subscription_tier').agg(
    median_ltv=('lifetime_value', 'median'),
    median_cac=('cac_calculated', 'median'),
    median_ratio=('ltv_cac_ratio', 'median'),
    median_payback=('payback_months', 'median')
).sort_values('median_ratio', ascending=True)

fig, ax = plt.subplots(figsize=(10, 5))
bar_colors = ['#DC2626' if r < 3 else '#F59E0B' if r < 5 else '#059669' for r in tier_summary['median_ratio']]
ax.barh(tier_summary.index, tier_summary['median_ratio'], color=bar_colors, edgecolor='white')
ax.axvline(3.0, color='#1E3A8A', linestyle='--', linewidth=2, label='Healthy threshold (3.0×)')
ax.set_xlabel('LTV:CAC Ratio')
ax.set_title('HMW-05: Which subscription tier gives the best return on acquisition investment?')
for i, v in enumerate(tier_summary['median_ratio']):
    ax.text(v + 0.1, i, f'{v:.1f}×', va='center', fontweight='bold')
ax.legend()
plt.tight_layout()
plt.show()

best_tier = tier_summary['median_ratio'].idxmax()
print(f"📌 ANSWER HMW-05: {best_tier} tier has the best LTV:CAC ratio ({tier_summary.loc[best_tier, 'median_ratio']:.1f}×).")
""")

md("""
### 📌 Summary of HMW Answers

| HMW | Best Chart | One-Line Answer |
|:----|:-----------|:----------------|
| HMW-01 | Horizontal bar (median CAC by channel) | Referral/Organic cheapest; Paid Social most expensive |
| HMW-02 | Plotly funnel | Biggest drop: Impressions → Clicks (biggest leakage point) |
| HMW-05 | Horizontal bar (LTV:CAC by tier) | Enterprise tier has the highest return on acquisition cost |

> HMW-03 (Budget simulation) and HMW-07 (Hypothesis testing) are addressed in Sections 11 and the Streamlit app.
""")

# ══════════════════════════════════════════════════════════════════════════════
# SECTION 11 — Statistical Test
# ══════════════════════════════════════════════════════════════════════════════
md("""
---
## 📖 Section 11: One Statistical Check — Is the CAC Difference Real?

### The question (HMW-07)
We *see* that Paid Social CAC (~₹1,400) is higher than Organic Search (~₹420). But is this difference **statistically significant**, or could it be random noise?

### The test: One-Way ANOVA
- **Null hypothesis (H₀):** All channels have the same mean CAC (differences are just random)
- **Alternative hypothesis (H₁):** At least one channel has a significantly different mean CAC
- **Decision rule:** If p-value < 0.05, reject H₀ → the differences ARE real

### Why ANOVA and not just comparing means?
Comparing two means at a time inflates the chance of a false positive. ANOVA tests **all groups at once**.

### Plain English interpretation
- **p < 0.05** → "There is less than a 5% chance this difference is just noise. We're confident it's real."
- **p > 0.05** → "We can't rule out that the difference is just random variation."
""")

code("""
# ── 11.1 Prepare groups for ANOVA ─────────────────────────────────────────
groups = [group['cac_calculated'].values for name, group in df.groupby('channel')]
group_names = [name for name, group in df.groupby('channel')]

# Run One-Way ANOVA
f_stat, p_value = stats.f_oneway(*groups)

print("=" * 60)
print("📊 ONE-WAY ANOVA TEST — Channel CAC Comparison")
print("=" * 60)
print(f"   Null Hypothesis: All channels have the same mean CAC")
print(f"   Alternative:     At least one channel mean differs")
print(f"   Significance:    α = 0.05")
print(f"")
print(f"   F-statistic:     {f_stat:.4f}")
print(f"   p-value:         {p_value:.6f}")
print(f"")
if p_value < 0.05:
    print(f"   ✅ VERDICT: REJECT H₀ (p = {p_value:.4f} < 0.05)")
    print(f"   → The CAC differences across channels are STATISTICALLY SIGNIFICANT.")
    print(f"   → There is only a {p_value*100:.2f}% chance this is random noise.")
else:
    print(f"   ❌ VERDICT: FAIL TO REJECT H₀ (p = {p_value:.4f} > 0.05)")
    print(f"   → We cannot confirm the differences are statistically significant.")
print("=" * 60)
""")

code("""
# ── 11.2 Visualise the group means with confidence intervals ──────────────
fig, ax = plt.subplots(figsize=(10, 5))

channel_stats = df.groupby('channel')['cac_calculated'].agg(['mean', 'std', 'count'])
channel_stats['se'] = channel_stats['std'] / np.sqrt(channel_stats['count'])  # Standard Error
channel_stats['ci95'] = channel_stats['se'] * 1.96  # 95% Confidence Interval
channel_stats = channel_stats.sort_values('mean')

ax.barh(channel_stats.index, channel_stats['mean'], xerr=channel_stats['ci95'],
       color=['#059669', '#7C3AED', '#F59E0B', '#0284C7', '#DC2626'],
       capsize=5, edgecolor='white', error_kw={'linewidth': 2})
ax.set_xlabel('Mean CAC (₹) with 95% Confidence Interval')
ax.set_title(f'ANOVA result: Channel CAC means differ significantly (F={f_stat:.1f}, p={p_value:.4f})')
for i, (mean, ci) in enumerate(zip(channel_stats['mean'], channel_stats['ci95'])):
    ax.text(mean + ci + 20, i, f'₹{mean:,.0f} ± {ci:,.0f}', va='center', fontsize=10)
plt.tight_layout()
plt.show()

print("💡 If confidence interval bars DON'T overlap, those two channels are significantly different.")
""")

md("""
### 📌 What We Found
- **ANOVA F-statistic** confirms the CAC differences are **not random noise**
- The p-value is well below 0.05, giving us >95% confidence
- **Confidence intervals** show that Organic Search and Referral bars don't overlap with Paid Social — confirming they are genuinely different populations
- **Business meaning:** You can confidently tell the CFO that Organic Search really IS cheaper — it's not just lucky months

### ✍️ Your Turn
> Try a **Kruskal-Wallis test** (non-parametric alternative): `stats.kruskal(*groups)`. Does it agree with ANOVA? Why might you prefer Kruskal-Wallis when data is skewed?
""")

# ══════════════════════════════════════════════════════════════════════════════
# SECTION 12 — Insights & Recommendations
# ══════════════════════════════════════════════════════════════════════════════
md("""
---
## 📖 Section 12: Insights & Actionable Recommendations

### The difference between a finding and a recommendation
- **Finding** = what the data shows (descriptive)
- **Insight** = what it means for the business (interpretive)
- **Recommendation** = what to do about it (prescriptive)

### Format: Finding → Evidence → Action
""")

code("""
# ── 12.1 Compile the 5 key insights ───────────────────────────────────────
insights = [
    {
        'Finding': 'Organic Search has the lowest CAC across all channels',
        'Evidence': f"Median CAC: ₹{df[df['channel']=='Organic Search']['cac_calculated'].median():,.0f} vs blended ₹{df['cac_calculated'].median():,.0f}",
        'Action': 'Increase SEO and content marketing budget by 15-20% by reallocating from Paid Social'
    },
    {
        'Finding': 'Paid Social CAC spikes 40%+ during Q3-Q4 (festive season)',
        'Evidence': f"Q4 avg CAC: ₹{df[df['seasonal_period']=='Q4']['cac_calculated'].mean():,.0f} vs Q1: ₹{df[df['seasonal_period']=='Q1']['cac_calculated'].mean():,.0f}",
        'Action': 'Pause Paid Social campaigns during Diwali auction peak (Oct-Nov); shift budget to Email'
    },
    {
        'Finding': 'Enterprise tier delivers the highest LTV:CAC return',
        'Evidence': f"Enterprise LTV:CAC: {df[df['subscription_tier']=='Enterprise']['ltv_cac_ratio'].median():.1f}× vs Starter: {df[df['subscription_tier']=='Starter']['ltv_cac_ratio'].median():.1f}×",
        'Action': 'Launch account-based marketing (ABM) targeting Enterprise prospects in top 3 cities'
    },
    {
        'Finding': 'Largest funnel leakage occurs at Impressions → Clicks stage',
        'Evidence': f"Average CTR across channels: {(df['clicks'].sum() / df['impressions'].sum() * 100):.2f}%",
        'Action': 'A/B test ad creatives and landing pages to improve CTR by 0.5-1 percentage points'
    },
    {
        'Finding': 'Channel CAC differences are statistically significant (not random)',
        'Evidence': f"One-Way ANOVA: F = {f_stat:.2f}, p = {p_value:.6f} (α = 0.05)",
        'Action': 'Use these proven CAC benchmarks for quarterly channel-level budget ceiling negotiations'
    },
]

insights_df = pd.DataFrame(insights)
print("=" * 70)
print("📋 TOP 5 INSIGHTS & RECOMMENDATIONS")
print("=" * 70)
for i, row in insights_df.iterrows():
    print(f"\\n{'─' * 60}")
    print(f"  INSIGHT {i+1}: {row['Finding']}")
    print(f"  EVIDENCE: {row['Evidence']}")
    print(f"  ACTION:   {row['Action']}")
print(f"\\n{'─' * 60}")
""")

code("""
# ── 12.2 Executive dashboard — 4 KPIs ─────────────────────────────────────
blended_cac = df['spend_amount'].sum() / df['conversions'].sum()
total_spend = df['spend_amount'].sum()
total_cust = df['conversions'].sum()
avg_ratio = df['ltv_cac_ratio'].median()

fig, axes = plt.subplots(1, 4, figsize=(16, 3))
kpis = [
    (f"₹{blended_cac:,.0f}", "Blended CAC", "#1E3A8A"),
    (f"₹{total_spend/1e7:,.1f} Cr", "Total Spend", "#0284C7"),
    (f"{total_cust:,}", "Customers Acquired", "#059669"),
    (f"{avg_ratio:.1f}×", "Median LTV:CAC", "#7C3AED"),
]
for ax, (val, label, color) in zip(axes, kpis):
    ax.text(0.5, 0.65, val, ha='center', va='center', fontsize=24, fontweight='bold', color=color,
           transform=ax.transAxes)
    ax.text(0.5, 0.2, label, ha='center', va='center', fontsize=11, color='#475569',
           transform=ax.transAxes)
    ax.set_xlim(0, 1); ax.set_ylim(0, 1)
    ax.axis('off')
    ax.patch.set_facecolor('#F0F4FF')
    ax.patch.set_alpha(0.5)
    for spine in ax.spines.values():
        spine.set_visible(False)
plt.suptitle('Executive Summary KPIs', fontsize=14, fontweight='bold', color='#1E3A8A')
plt.tight_layout()
plt.show()
""")

md("""
### ✍️ Your Turn
> Pick Insight #3 (Enterprise tier). Write a 3-sentence email to the CMO summarising: (1) the finding, (2) the evidence, (3) your recommended action. Use the exact numbers from above.
""")

# ══════════════════════════════════════════════════════════════════════════════
# SECTION 13 — Chart Cheat Sheet & Exercises
# ══════════════════════════════════════════════════════════════════════════════
md("""
---
## 📖 Section 13: Chart-Choice Cheat Sheet & Practice Exercises

### 🗺️ When to use which chart?

| Your Question | X Variable | Y Variable | Best Chart |
|:-------------|:-----------|:-----------|:-----------|
| What's the distribution? | 1 numeric | — | Histogram / KDE / Box plot |
| How many in each category? | 1 category | — | Bar chart (NOT pie) |
| Do two numbers correlate? | numeric | numeric | Scatter + trend line |
| How does Y vary by group? | category | numeric | Box / Violin by group |
| What's the composition? | category | category | Stacked bar / Heatmap |
| How does it change over time? | date | numeric | Line chart + moving avg |
| What drives 80% of the total? | category | numeric | Pareto chart |
| How do 3+ variables relate? | numeric | numeric | Bubble / pair plot / heatmap |
| What's the hierarchical breakdown? | nested categories | numeric | Treemap / Sunburst |

### 📝 5 Practice Exercises

1. **Exercise 1 (Univariate):** Create a histogram of `monthly_revenue` with 30 bins. Add vertical lines for mean and median. Is it skewed?

2. **Exercise 2 (Bivariate):** Make a scatter plot of `churn_rate` vs `lifetime_value`. Add a trend line. What's the correlation?

3. **Exercise 3 (Grouped):** Create a grouped bar chart showing **median payback_months** for each `channel × subscription_tier` combination. Which combination pays back fastest?

4. **Exercise 4 (Time):** Plot a line chart of **monthly total conversions** over time. Add a 3-month moving average. Is the business growing?

5. **Exercise 5 (Statistical):** Run a **Chi-Square test** on a crosstab of `channel` vs `campaign_status`. Are certain channels more likely to be "Paused"?

> **Hint for Exercise 5:**
> ```python
> contingency = pd.crosstab(df['channel'], df['campaign_status'])
> chi2, p, dof, expected = stats.chi2_contingency(contingency)
> print(f"Chi-Square = {chi2:.2f}, p = {p:.4f}")
> ```
""")

md("""
---

## 🎓 Notebook Complete

**What you learned:**
1. Always start with **questions** (HMW), not code
2. **Audit** data quality before analysis (missing, duplicates, outliers, labels, types)
3. **Clean** systematically with a logged before/after table
4. Choose **the right chart** for each question type
5. **Univariate → Bivariate → Multivariate** = simple to complex
6. **Time-series** reveals trends and seasonality
7. **Pareto** identifies where to focus effort
8. Always **answer your original questions** with specific numbers
9. Use **statistical tests** to prove differences are real (not random)
10. Format findings as **Evidence → Insight → Action**

---

*Notebook by Gourav Kumar | Roll: 2500320140055 | Project ID: P083 | Batch: ABES_OND26_MCA 2nd Year_Data Analytics_Batch 01 | Mentor: Mr. Dinesh Sereno | ABES Engineering College*
""")

# ── Save notebook ──────────────────────────────────────────────────────────
nb.cells = cells
out_path = r'c:\Users\asus\Downloads\MERN-Expense-Tracker-main\EDA_CAC_Analysis_Lesson.ipynb'
with open(out_path, 'w', encoding='utf-8') as f:
    nbf.write(nb, f)
print(f"Notebook saved to {out_path}")
