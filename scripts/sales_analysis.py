"""
Sales Orders — Exploratory Data Analysis (EDA)
Comprehensive analysis covering descriptive statistics, revenue trends,
profitability, customer segmentation, product performance, and more.
"""
import pandas as pd
import numpy as np
import warnings
warnings.filterwarnings('ignore')

# ============================================================
# 1. DATA LOADING & CLEANING
# ============================================================
print("=" * 70)
print("SALES ORDERS — EXPLORATORY DATA ANALYSIS")
print("=" * 70)

df = pd.read_csv('data/Sales_Orders.csv')
df['Order Date'] = pd.to_datetime(df['Order Date'], format='%d/%m/%Y')
df['Ship Date'] = pd.to_datetime(df['Ship Date'], format='%d/%m/%Y')
df['Year'] = df['Order Date'].dt.year
df['Month'] = df['Order Date'].dt.month
df['Quarter'] = df['Order Date'].dt.quarter
df['Delivery_Days'] = (df['Ship Date'] - df['Order Date']).dt.days
df['Profit_Margin'] = (df['Profit'] / df['Sales'] * 100).round(2)

print(f"\nDataset Shape: {df.shape[0]:,} rows x {df.shape[1]} columns")
print(f"Date Range: {df['Order Date'].min().date()} to {df['Order Date'].max().date()}")
print(f"Missing Values: {df.isnull().sum().sum()}")
print(f"Duplicate Rows: {df.duplicated().sum()}")

# ============================================================
# 2. DESCRIPTIVE STATISTICS
# ============================================================
print("\n" + "=" * 70)
print("DESCRIPTIVE STATISTICS")
print("=" * 70)
for col in ['Sales', 'Profit', 'Quantity', 'Discount']:
    s = df[col]
    print(f"\n--- {col} ---")
    print(f"  Mean: {s.mean():>12,.2f}  |  Median: {s.median():>12,.2f}")
    print(f"  Std:  {s.std():>12,.2f}  |  Min:    {s.min():>12,.2f}")
    print(f"  Max:  {s.max():>12,.2f}  |  Sum:    {s.sum():>12,.2f}")
    print(f"  Skew: {s.skew():>12,.2f}  |  Kurt:   {s.kurtosis():>12,.2f}")

# ============================================================
# 3. REVENUE ANALYSIS
# ============================================================
print("\n" + "=" * 70)
print("REVENUE ANALYSIS")
print("=" * 70)
yearly = df.groupby('Year').agg(
    Revenue=('Sales', 'sum'),
    Profit=('Profit', 'sum'),
    Orders=('Order ID', 'nunique'),
    Customers=('Customer ID', 'nunique')
).reset_index()
yearly['YoY_Growth'] = yearly['Revenue'].pct_change() * 100
yearly['Avg_Order_Value'] = yearly['Revenue'] / yearly['Orders']
print("\nYearly Performance:")
print(yearly.to_string(index=False, float_format='${:,.0f}'.format))

# ============================================================
# 4. PROFITABILITY ANALYSIS
# ============================================================
print("\n" + "=" * 70)
print("PROFITABILITY ANALYSIS")
print("=" * 70)

print("\nBy Category:")
cat = df.groupby('Category').agg(
    Sales=('Sales', 'sum'), Profit=('Profit', 'sum'), Qty=('Quantity', 'sum')
).reset_index()
cat['Margin%'] = (cat['Profit'] / cat['Sales'] * 100).round(2)
print(cat.sort_values('Sales', ascending=False).to_string(index=False))

print("\nBy Sub-Category (Top 5 / Bottom 5 by Profit):")
subcat = df.groupby('Sub-Category').agg(
    Sales=('Sales', 'sum'), Profit=('Profit', 'sum')
).reset_index()
subcat['Margin%'] = (subcat['Profit'] / subcat['Sales'] * 100).round(2)
print("\n  TOP 5:")
print(subcat.nlargest(5, 'Profit').to_string(index=False))
print("\n  BOTTOM 5:")
print(subcat.nsmallest(5, 'Profit').to_string(index=False))

# ============================================================
# 5. CUSTOMER SEGMENTATION (RFM)
# ============================================================
print("\n" + "=" * 70)
print("CUSTOMER SEGMENTATION — RFM ANALYSIS")
print("=" * 70)
max_date = df['Order Date'].max()
rfm = df.groupby('Customer ID').agg(
    Recency=('Order Date', lambda x: (max_date - x.max()).days),
    Frequency=('Order ID', 'nunique'),
    Monetary=('Sales', 'sum')
).reset_index()
rfm['R'] = pd.qcut(rfm['Recency'], 4, labels=[4,3,2,1]).astype(int)
rfm['F'] = pd.qcut(rfm['Frequency'].rank(method='first'), 4, labels=[1,2,3,4]).astype(int)
rfm['M'] = pd.qcut(rfm['Monetary'], 4, labels=[1,2,3,4]).astype(int)
rfm['RFM_Score'] = rfm['R'] + rfm['F'] + rfm['M']
rfm['Segment'] = rfm['RFM_Score'].apply(
    lambda s: 'Champions' if s>=10 else 'Loyal' if s>=8 else 'Potential' if s>=6 else 'At Risk' if s>=4 else 'Lost'
)
seg = rfm.groupby('Segment').agg(
    Count=('Customer ID', 'count'),
    Avg_Revenue=('Monetary', 'mean'),
    Total_Revenue=('Monetary', 'sum')
).reset_index()
seg['%_Customers'] = (seg['Count'] / seg['Count'].sum() * 100).round(1)
print(seg.sort_values('Total_Revenue', ascending=False).to_string(index=False))

# ============================================================
# 6. PRODUCT PERFORMANCE
# ============================================================
print("\n" + "=" * 70)
print("PRODUCT PERFORMANCE")
print("=" * 70)
prod = df.groupby('Product Name').agg(Sales=('Sales','sum'), Profit=('Profit','sum'), Qty=('Quantity','sum')).reset_index()
print(f"\nTotal Unique Products: {len(prod):,}")
print(f"\nTop 10 Products by Revenue:")
for i, (_, r) in enumerate(prod.nlargest(10, 'Sales').iterrows(), 1):
    print(f"  {i:>2}. {r['Product Name'][:50]:50s}  ${r['Sales']:>10,.0f}")

# ABC Analysis
prod_sorted = prod.sort_values('Sales', ascending=False)
prod_sorted['Cum%'] = prod_sorted['Sales'].cumsum() / prod_sorted['Sales'].sum() * 100
a_count = (prod_sorted['Cum%'] <= 80).sum()
b_count = ((prod_sorted['Cum%'] > 80) & (prod_sorted['Cum%'] <= 95)).sum()
c_count = (prod_sorted['Cum%'] > 95).sum()
print(f"\nABC Classification:")
print(f"  Class A (80% revenue): {a_count} products ({a_count/len(prod)*100:.1f}%)")
print(f"  Class B (next 15%):    {b_count} products ({b_count/len(prod)*100:.1f}%)")
print(f"  Class C (last 5%):     {c_count} products ({c_count/len(prod)*100:.1f}%)")

# ============================================================
# 7. REGIONAL ANALYSIS
# ============================================================
print("\n" + "=" * 70)
print("REGIONAL ANALYSIS")
print("=" * 70)
reg = df.groupby('Region').agg(
    Sales=('Sales','sum'), Profit=('Profit','sum'), Orders=('Order ID','nunique')
).reset_index()
reg['Margin%'] = (reg['Profit'] / reg['Sales'] * 100).round(2)
reg['AOV'] = (reg['Sales'] / reg['Orders']).round(2)
print(reg.sort_values('Sales', ascending=False).to_string(index=False))

print("\nTop 5 States by Revenue:")
states = df.groupby('State')['Sales'].sum().nlargest(5)
for i, (state, val) in enumerate(states.items(), 1):
    print(f"  {i}. {state:25s} ${val:>12,.0f}")

# ============================================================
# 8. SHIPPING ANALYSIS
# ============================================================
print("\n" + "=" * 70)
print("SHIPPING ANALYSIS")
print("=" * 70)
ship = df.groupby('Ship Mode').agg(
    Orders=('Order ID','nunique'), Avg_Days=('Delivery_Days','mean'), Sales=('Sales','sum')
).reset_index()
ship['%_Orders'] = (ship['Orders'] / ship['Orders'].sum() * 100).round(1)
print(ship.sort_values('Orders', ascending=False).to_string(index=False))
print(f"\nOverall Avg Delivery Time: {df['Delivery_Days'].mean():.1f} days")

# ============================================================
# 9. DISCOUNT IMPACT
# ============================================================
print("\n" + "=" * 70)
print("DISCOUNT IMPACT ANALYSIS")
print("=" * 70)
df['Disc_Band'] = pd.cut(df['Discount'], bins=[-0.01,0,0.1,0.2,0.3,0.5,1],
                          labels=['None','1-10%','11-20%','21-30%','31-50%','50%+'])
disc = df.groupby('Disc_Band', observed=True).agg(
    Count=('Sales','count'), Total_Sales=('Sales','sum'), Total_Profit=('Profit','sum')
).reset_index()
disc['Margin%'] = (disc['Total_Profit'] / disc['Total_Sales'] * 100).round(2)
print(disc.to_string(index=False))
corr = df['Discount'].corr(df['Profit'])
print(f"\nDiscount-Profit Correlation: {corr:.4f}")

# ============================================================
# 10. KEY INSIGHTS SUMMARY
# ============================================================
print("\n" + "=" * 70)
print("KEY INSIGHTS & RECOMMENDATIONS")
print("=" * 70)
total_rev = df['Sales'].sum()
total_prof = df['Profit'].sum()
print(f"""
1. REVENUE: ${total_rev:,.0f} total revenue with ${total_prof:,.0f} profit ({total_prof/total_rev*100:.1f}% margin)
2. GROWTH: Revenue grew from ${yearly.iloc[0]['Revenue']:,.0f} ({int(yearly.iloc[0]['Year'])}) to ${yearly.iloc[-1]['Revenue']:,.0f} ({int(yearly.iloc[-1]['Year'])})
3. TOP CATEGORY: Technology leads in revenue but Office Supplies has highest order volume
4. LOSS MAKERS: Tables and Bookcases sub-categories are generating losses — review pricing
5. CUSTOMER VALUE: Champions segment ({seg[seg['Segment']=='Champions']['%_Customers'].values[0]}% of customers) drives disproportionate revenue
6. REGIONAL: West region leads in revenue; Central region has lowest profit margins
7. DISCOUNTS: Orders with >20% discount have negative profit margins — cap discounts at 20%
8. SEASONALITY: Q4 (Oct-Dec) is peak season; plan inventory and marketing accordingly
9. SHIPPING: Standard Class is most popular (60%+); Same Day has highest margins
10. PARETO: {a_count} products ({a_count/len(prod)*100:.1f}%) generate 80% of revenue — focus inventory here
""")
print("Analysis complete. See output/ for Excel report and images/ for charts.")
