"""
Sales Analytics - Visualization Generator
Generates publication-quality charts for all analysis dimensions.
"""
import pandas as pd
import numpy as np
import matplotlib
matplotlib.use('Agg')
import matplotlib.pyplot as plt
import matplotlib.ticker as mticker
import seaborn as sns
from matplotlib.gridspec import GridSpec
import warnings
warnings.filterwarnings('ignore')

sns.set_theme(style="whitegrid", palette="muted")
plt.rcParams.update({
    'figure.dpi': 150, 'savefig.dpi': 150, 'font.family': 'sans-serif',
    'font.size': 10, 'axes.titlesize': 13, 'axes.labelsize': 11,
    'figure.facecolor': 'white', 'axes.facecolor': '#FAFAFA',
    'grid.alpha': 0.3
})
COLORS = ['#2E86AB', '#A23B72', '#F18F01', '#C73E1D', '#3B1F2B', '#44BBA4', '#E94F37', '#393E41']
IMG_DIR = 'images'

df = pd.read_csv('data/Sales_Orders.csv')
df['Order Date'] = pd.to_datetime(df['Order Date'], format='%d/%m/%Y')
df['Ship Date'] = pd.to_datetime(df['Ship Date'], format='%d/%m/%Y')
df['Year'] = df['Order Date'].dt.year
df['Month'] = df['Order Date'].dt.month
df['Quarter'] = df['Order Date'].dt.quarter
df['YearMonth'] = df['Order Date'].dt.to_period('M')
df['Delivery_Days'] = (df['Ship Date'] - df['Order Date']).dt.days
df['Profit_Margin'] = df['Profit'] / df['Sales'] * 100

def currency_fmt(x, _): return f'${x:,.0f}'
def save(fig, name): fig.savefig(f'{IMG_DIR}/{name}.png', bbox_inches='tight', pad_inches=0.3); plt.close(fig)

# ====== 1. EXECUTIVE KPI DASHBOARD ======
fig = plt.figure(figsize=(16, 10))
fig.suptitle('Sales Analytics — Executive Dashboard', fontsize=18, fontweight='bold', color='#1a1a2e', y=0.98)
gs = GridSpec(3, 4, figure=fig, hspace=0.45, wspace=0.35)

kpis = [
    ('Total Revenue', f"${df['Sales'].sum():,.0f}"),
    ('Total Profit', f"${df['Profit'].sum():,.0f}"),
    ('Profit Margin', f"{df['Profit'].sum()/df['Sales'].sum()*100:.1f}%"),
    ('Total Orders', f"{df['Order ID'].nunique():,}"),
]
for i, (label, val) in enumerate(kpis):
    ax = fig.add_subplot(gs[0, i])
    ax.text(0.5, 0.6, val, ha='center', va='center', fontsize=20, fontweight='bold', color=COLORS[i], transform=ax.transAxes)
    ax.text(0.5, 0.2, label, ha='center', va='center', fontsize=11, color='#555', transform=ax.transAxes)
    ax.set_xlim(0,1); ax.set_ylim(0,1); ax.axis('off')
    ax.patch.set_facecolor('#f0f4f8'); ax.patch.set_alpha(0.8)
    for spine in ax.spines.values(): spine.set_visible(True); spine.set_color('#ddd')

ax1 = fig.add_subplot(gs[1, :2])
yearly = df.groupby('Year')['Sales'].sum()
bars = ax1.bar(yearly.index.astype(str), yearly.values, color=COLORS[:4], edgecolor='white', linewidth=0.5)
ax1.set_title('Annual Revenue', fontweight='bold')
ax1.yaxis.set_major_formatter(mticker.FuncFormatter(currency_fmt))
for bar, val in zip(bars, yearly.values):
    ax1.text(bar.get_x() + bar.get_width()/2, bar.get_height() + 5000, f'${val:,.0f}', ha='center', va='bottom', fontsize=9, fontweight='bold')

ax2 = fig.add_subplot(gs[1, 2:])
cat = df.groupby('Category')['Sales'].sum()
wedges, texts, autotexts = ax2.pie(cat.values, labels=cat.index, autopct='%1.1f%%', colors=COLORS[:3], startangle=90, textprops={'fontsize': 10})
ax2.set_title('Revenue by Category', fontweight='bold')

ax3 = fig.add_subplot(gs[2, :2])
monthly_rev = df.groupby(df['Order Date'].dt.to_period('M'))['Sales'].sum()
ax3.plot(range(len(monthly_rev)), monthly_rev.values, color=COLORS[0], linewidth=1.5)
ax3.fill_between(range(len(monthly_rev)), monthly_rev.values, alpha=0.15, color=COLORS[0])
ax3.set_title('Monthly Revenue Trend', fontweight='bold')
ax3.yaxis.set_major_formatter(mticker.FuncFormatter(currency_fmt))
tick_positions = list(range(0, len(monthly_rev), 6))
ax3.set_xticks(tick_positions)
ax3.set_xticklabels([str(monthly_rev.index[i]) for i in tick_positions], rotation=45, fontsize=8)

ax4 = fig.add_subplot(gs[2, 2:])
region = df.groupby('Region')[['Sales','Profit']].sum().sort_values('Sales', ascending=True)
y_pos = range(len(region))
ax4.barh(y_pos, region['Sales'], color=COLORS[0], alpha=0.8, label='Sales', height=0.4)
ax4.barh([y+0.4 for y in y_pos], region['Profit'], color=COLORS[1], alpha=0.8, label='Profit', height=0.4)
ax4.set_yticks([y+0.2 for y in y_pos]); ax4.set_yticklabels(region.index)
ax4.set_title('Sales & Profit by Region', fontweight='bold')
ax4.xaxis.set_major_formatter(mticker.FuncFormatter(currency_fmt))
ax4.legend(fontsize=9)
save(fig, '01_executive_dashboard')

# ====== 2. REVENUE ANALYSIS ======
fig, axes = plt.subplots(2, 2, figsize=(16, 10))
fig.suptitle('Revenue Analysis', fontsize=16, fontweight='bold', y=1.01)

yearly_rev = df.groupby('Year').agg({'Sales': 'sum', 'Profit': 'sum'}).reset_index()
x = np.arange(len(yearly_rev))
axes[0,0].bar(x-0.2, yearly_rev['Sales'], 0.4, label='Revenue', color=COLORS[0])
axes[0,0].bar(x+0.2, yearly_rev['Profit'], 0.4, label='Profit', color=COLORS[1])
axes[0,0].set_xticks(x); axes[0,0].set_xticklabels(yearly_rev['Year'].astype(int))
axes[0,0].set_title('Yearly Revenue vs Profit'); axes[0,0].legend()
axes[0,0].yaxis.set_major_formatter(mticker.FuncFormatter(currency_fmt))

for year in sorted(df['Year'].unique()):
    monthly = df[df['Year']==year].groupby('Month')['Sales'].sum()
    axes[0,1].plot(monthly.index, monthly.values, marker='o', markersize=4, label=str(int(year)))
axes[0,1].set_title('Monthly Revenue by Year'); axes[0,1].legend(fontsize=8)
axes[0,1].set_xticks(range(1,13))
axes[0,1].set_xticklabels(['J','F','M','A','M','J','J','A','S','O','N','D'])
axes[0,1].yaxis.set_major_formatter(mticker.FuncFormatter(currency_fmt))

quarterly = df.groupby([df['Year'], df['Quarter']])['Sales'].sum().reset_index()
quarterly['Label'] = quarterly['Year'].astype(int).astype(str) + ' Q' + quarterly['Quarter'].astype(str)
axes[1,0].bar(range(len(quarterly)), quarterly['Sales'], color=[COLORS[q-1] for q in quarterly['Quarter']])
axes[1,0].set_xticks(range(len(quarterly)))
axes[1,0].set_xticklabels(quarterly['Label'], rotation=45, fontsize=7)
axes[1,0].set_title('Quarterly Revenue')
axes[1,0].yaxis.set_major_formatter(mticker.FuncFormatter(currency_fmt))

yoy = yearly_rev.copy()
yoy['Growth'] = yoy['Sales'].pct_change() * 100
yoy = yoy.dropna()
colors_yoy = [COLORS[0] if g > 0 else COLORS[3] for g in yoy['Growth']]
axes[1,1].bar(yoy['Year'].astype(int).astype(str), yoy['Growth'], color=colors_yoy)
axes[1,1].set_title('Year-over-Year Revenue Growth (%)')
axes[1,1].axhline(y=0, color='gray', linestyle='--', linewidth=0.5)
for i, (_, row) in enumerate(yoy.iterrows()):
    axes[1,1].text(i, row['Growth']+0.5, f"{row['Growth']:.1f}%", ha='center', fontsize=10, fontweight='bold')

plt.tight_layout()
save(fig, '02_revenue_analysis')

# ====== 3. PROFITABILITY ANALYSIS ======
fig, axes = plt.subplots(2, 2, figsize=(16, 10))
fig.suptitle('Profitability Analysis', fontsize=16, fontweight='bold', y=1.01)

subcat = df.groupby('Sub-Category').agg({'Profit': 'sum'}).sort_values('Profit')
colors_p = [COLORS[3] if v < 0 else COLORS[0] for v in subcat['Profit']]
axes[0,0].barh(subcat.index, subcat['Profit'], color=colors_p)
axes[0,0].set_title('Profit by Sub-Category')
axes[0,0].axvline(x=0, color='black', linewidth=0.5)
axes[0,0].xaxis.set_major_formatter(mticker.FuncFormatter(currency_fmt))

cat_margin = df.groupby('Category').apply(lambda x: x['Profit'].sum()/x['Sales'].sum()*100).sort_values()
axes[0,1].barh(cat_margin.index, cat_margin.values, color=COLORS[:3])
axes[0,1].set_title('Profit Margin by Category (%)')
for i, v in enumerate(cat_margin.values):
    axes[0,1].text(v+0.3, i, f'{v:.1f}%', va='center', fontweight='bold')

seg_prof = df.groupby('Segment').agg({'Sales': 'sum', 'Profit': 'sum'}).reset_index()
seg_prof['Margin'] = seg_prof['Profit'] / seg_prof['Sales'] * 100
axes[1,0].bar(seg_prof['Segment'], seg_prof['Margin'], color=COLORS[:3])
axes[1,0].set_title('Profit Margin by Segment (%)')
for i, v in enumerate(seg_prof['Margin']):
    axes[1,0].text(i, v+0.2, f'{v:.1f}%', ha='center', fontweight='bold')

reg_prof = df.groupby('Region').apply(lambda x: x['Profit'].sum()/x['Sales'].sum()*100).sort_values()
axes[1,1].barh(reg_prof.index, reg_prof.values, color=COLORS[:4])
axes[1,1].set_title('Profit Margin by Region (%)')
for i, v in enumerate(reg_prof.values):
    axes[1,1].text(v+0.2, i, f'{v:.1f}%', va='center', fontweight='bold')

plt.tight_layout()
save(fig, '03_profitability_analysis')

# ====== 4. CUSTOMER SEGMENTATION (RFM) ======
max_date = df['Order Date'].max()
rfm = df.groupby('Customer ID').agg({
    'Order Date': lambda x: (max_date - x.max()).days,
    'Order ID': 'nunique',
    'Sales': 'sum'
}).reset_index()
rfm.columns = ['Customer ID', 'Recency', 'Frequency', 'Monetary']
rfm['R'] = pd.qcut(rfm['Recency'], 4, labels=[4,3,2,1]).astype(int)
rfm['F'] = pd.qcut(rfm['Frequency'].rank(method='first'), 4, labels=[1,2,3,4]).astype(int)
rfm['M'] = pd.qcut(rfm['Monetary'], 4, labels=[1,2,3,4]).astype(int)
rfm['Score'] = rfm['R'] + rfm['F'] + rfm['M']
rfm['Segment'] = rfm['Score'].apply(lambda s: 'Champions' if s>=10 else 'Loyal' if s>=8 else 'Potential' if s>=6 else 'At Risk' if s>=4 else 'Lost')

fig, axes = plt.subplots(2, 2, figsize=(16, 10))
fig.suptitle('Customer Segmentation — RFM Analysis', fontsize=16, fontweight='bold', y=1.01)

seg_counts = rfm['Segment'].value_counts()
axes[0,0].pie(seg_counts, labels=seg_counts.index, autopct='%1.1f%%', colors=COLORS[:len(seg_counts)], startangle=90)
axes[0,0].set_title('Customer Segments Distribution')

seg_rev = rfm.groupby('Segment')['Monetary'].sum().sort_values(ascending=True)
axes[0,1].barh(seg_rev.index, seg_rev.values, color=COLORS[:len(seg_rev)])
axes[0,1].set_title('Revenue by Customer Segment')
axes[0,1].xaxis.set_major_formatter(mticker.FuncFormatter(currency_fmt))

axes[1,0].scatter(rfm['Recency'], rfm['Monetary'], c=rfm['Score'], cmap='RdYlGn', alpha=0.5, s=20)
axes[1,0].set_xlabel('Recency (days)'); axes[1,0].set_ylabel('Monetary ($)')
axes[1,0].set_title('Recency vs Monetary (colored by RFM Score)')

top20 = rfm.nlargest(20, 'Monetary')
axes[1,1].barh(range(20), top20['Monetary'].values, color=COLORS[0])
axes[1,1].set_yticks(range(20))
axes[1,1].set_yticklabels(top20['Customer ID'].values, fontsize=7)
axes[1,1].set_title('Top 20 Customers by Revenue')
axes[1,1].xaxis.set_major_formatter(mticker.FuncFormatter(currency_fmt))

plt.tight_layout()
save(fig, '04_customer_segmentation')

# ====== 5. PRODUCT PERFORMANCE ======
fig, axes = plt.subplots(2, 2, figsize=(16, 12))
fig.suptitle('Product Performance Analysis', fontsize=16, fontweight='bold', y=1.01)

top10 = df.groupby('Product Name')['Sales'].sum().nlargest(10)
axes[0,0].barh(range(10), top10.values, color=COLORS[0])
axes[0,0].set_yticks(range(10))
axes[0,0].set_yticklabels([n[:35]+'...' if len(n)>35 else n for n in top10.index], fontsize=8)
axes[0,0].set_title('Top 10 Products by Revenue')
axes[0,0].xaxis.set_major_formatter(mticker.FuncFormatter(currency_fmt))

bottom10 = df.groupby('Product Name')['Profit'].sum().nsmallest(10)
axes[0,1].barh(range(10), bottom10.values, color=COLORS[3])
axes[0,1].set_yticks(range(10))
axes[0,1].set_yticklabels([n[:35]+'...' if len(n)>35 else n for n in bottom10.index], fontsize=8)
axes[0,1].set_title('Top 10 Loss-Making Products')
axes[0,1].xaxis.set_major_formatter(mticker.FuncFormatter(currency_fmt))

# ABC / Pareto
prod_sales = df.groupby('Product Name')['Sales'].sum().sort_values(ascending=False)
cumulative = prod_sales.cumsum() / prod_sales.sum() * 100
ax_pareto = axes[1,0]
ax_pareto.bar(range(len(prod_sales)), prod_sales.values, color=COLORS[0], alpha=0.4, width=1.0)
ax2_p = ax_pareto.twinx()
ax2_p.plot(range(len(cumulative)), cumulative.values, color=COLORS[3], linewidth=2)
ax2_p.axhline(y=80, color='red', linestyle='--', alpha=0.5, label='80% threshold')
ax2_p.set_ylabel('Cumulative %')
ax_pareto.set_title('Pareto Analysis (ABC Classification)')
ax_pareto.set_xlabel('Products (sorted by revenue)')
ax_pareto.yaxis.set_major_formatter(mticker.FuncFormatter(currency_fmt))

subcat_qty = df.groupby('Sub-Category')['Quantity'].sum().sort_values(ascending=True)
axes[1,1].barh(subcat_qty.index, subcat_qty.values, color=COLORS[5])
axes[1,1].set_title('Units Sold by Sub-Category')

plt.tight_layout()
save(fig, '05_product_performance')

# ====== 6. REGIONAL ANALYSIS ======
fig, axes = plt.subplots(2, 2, figsize=(16, 10))
fig.suptitle('Regional Analysis', fontsize=16, fontweight='bold', y=1.01)

state_sales = df.groupby('State')['Sales'].sum().nlargest(10).sort_values()
axes[0,0].barh(state_sales.index, state_sales.values, color=COLORS[0])
axes[0,0].set_title('Top 10 States by Revenue')
axes[0,0].xaxis.set_major_formatter(mticker.FuncFormatter(currency_fmt))

state_profit = df.groupby('State')['Profit'].sum().nsmallest(10).sort_values()
colors_sp = [COLORS[3] if v<0 else COLORS[0] for v in state_profit.values]
axes[0,1].barh(state_profit.index, state_profit.values, color=colors_sp)
axes[0,1].set_title('Bottom 10 States by Profit')
axes[0,1].axvline(x=0, color='black', linewidth=0.5)
axes[0,1].xaxis.set_major_formatter(mticker.FuncFormatter(currency_fmt))

reg_cat = df.pivot_table(values='Sales', index='Region', columns='Category', aggfunc='sum')
reg_cat.plot(kind='bar', ax=axes[1,0], color=COLORS[:3])
axes[1,0].set_title('Revenue by Region & Category')
axes[1,0].yaxis.set_major_formatter(mticker.FuncFormatter(currency_fmt))
axes[1,0].tick_params(axis='x', rotation=0)
axes[1,0].legend(fontsize=9)

city_sales = df.groupby('City')['Sales'].sum().nlargest(10).sort_values()
axes[1,1].barh(city_sales.index, city_sales.values, color=COLORS[1])
axes[1,1].set_title('Top 10 Cities by Revenue')
axes[1,1].xaxis.set_major_formatter(mticker.FuncFormatter(currency_fmt))

plt.tight_layout()
save(fig, '06_regional_analysis')

# ====== 7. TIME SERIES & SEASONALITY ======
fig, axes = plt.subplots(2, 2, figsize=(16, 10))
fig.suptitle('Time Series & Seasonality Analysis', fontsize=16, fontweight='bold', y=1.01)

monthly_ts = df.groupby(df['Order Date'].dt.to_period('M'))['Sales'].sum()
x_range = range(len(monthly_ts))
axes[0,0].plot(x_range, monthly_ts.values, color=COLORS[0], alpha=0.4, linewidth=1)
ma3 = monthly_ts.rolling(3).mean()
ma6 = monthly_ts.rolling(6).mean()
axes[0,0].plot(x_range, ma3.values, color=COLORS[1], linewidth=2, label='3-Month MA')
axes[0,0].plot(x_range, ma6.values, color=COLORS[3], linewidth=2, label='6-Month MA')
axes[0,0].set_title('Revenue with Moving Averages')
axes[0,0].legend(fontsize=9)
axes[0,0].yaxis.set_major_formatter(mticker.FuncFormatter(currency_fmt))
ticks = list(range(0, len(monthly_ts), 6))
axes[0,0].set_xticks(ticks)
axes[0,0].set_xticklabels([str(monthly_ts.index[i]) for i in ticks], rotation=45, fontsize=7)

monthly_avg = df.groupby('Month')['Sales'].mean()
months = ['Jan','Feb','Mar','Apr','May','Jun','Jul','Aug','Sep','Oct','Nov','Dec']
avg_all = monthly_avg.mean()
colors_s = [COLORS[0] if v >= avg_all else COLORS[3] for v in monthly_avg.values]
axes[0,1].bar(months, monthly_avg.values, color=colors_s)
axes[0,1].axhline(y=avg_all, color='red', linestyle='--', alpha=0.7, label='Average')
axes[0,1].set_title('Seasonality — Avg Monthly Revenue')
axes[0,1].legend(fontsize=9)
axes[0,1].yaxis.set_major_formatter(mticker.FuncFormatter(currency_fmt))
axes[0,1].tick_params(axis='x', rotation=45)

dow_map = {0:'Mon',1:'Tue',2:'Wed',3:'Thu',4:'Fri',5:'Sat',6:'Sun'}
df['DOW'] = df['Order Date'].dt.dayofweek
dow_sales = df.groupby('DOW')['Sales'].mean()
axes[1,0].bar([dow_map[d] for d in dow_sales.index], dow_sales.values, color=COLORS[5])
axes[1,0].set_title('Avg Revenue by Day of Week')
axes[1,0].yaxis.set_major_formatter(mticker.FuncFormatter(currency_fmt))

qtr_sales = df.groupby('Quarter')['Sales'].mean()
axes[1,1].bar(['Q1','Q2','Q3','Q4'], qtr_sales.values, color=COLORS[:4])
axes[1,1].set_title('Avg Revenue by Quarter')
axes[1,1].yaxis.set_major_formatter(mticker.FuncFormatter(currency_fmt))
for i, v in enumerate(qtr_sales.values):
    axes[1,1].text(i, v+500, f'${v:,.0f}', ha='center', fontsize=9, fontweight='bold')

plt.tight_layout()
save(fig, '07_time_series_seasonality')

# ====== 8. SHIPPING ANALYSIS ======
fig, axes = plt.subplots(2, 2, figsize=(16, 10))
fig.suptitle('Shipping & Delivery Analysis', fontsize=16, fontweight='bold', y=1.01)

ship = df.groupby('Ship Mode').agg({'Order ID': 'nunique', 'Sales': 'sum'}).reset_index()
ship.columns = ['Ship Mode', 'Orders', 'Sales']
axes[0,0].pie(ship['Orders'], labels=ship['Ship Mode'], autopct='%1.1f%%', colors=COLORS[:4], startangle=90)
axes[0,0].set_title('Orders by Shipping Mode')

ship_del = df.groupby('Ship Mode')['Delivery_Days'].mean().sort_values()
axes[0,1].barh(ship_del.index, ship_del.values, color=COLORS[:4])
axes[0,1].set_title('Avg Delivery Days by Ship Mode')
for i, v in enumerate(ship_del.values):
    axes[0,1].text(v+0.1, i, f'{v:.1f} days', va='center', fontweight='bold')

axes[1,0].hist(df['Delivery_Days'], bins=range(0, df['Delivery_Days'].max()+2), color=COLORS[0], edgecolor='white', alpha=0.8)
axes[1,0].set_title('Delivery Time Distribution')
axes[1,0].set_xlabel('Days to Deliver')
axes[1,0].set_ylabel('Order Count')

ship_prof = df.groupby('Ship Mode').agg({'Profit': 'sum'}).sort_values('Profit')
colors_sh = [COLORS[3] if v<0 else COLORS[0] for v in ship_prof['Profit']]
axes[1,1].barh(ship_prof.index, ship_prof['Profit'], color=colors_sh)
axes[1,1].set_title('Profit by Shipping Mode')
axes[1,1].xaxis.set_major_formatter(mticker.FuncFormatter(currency_fmt))

plt.tight_layout()
save(fig, '08_shipping_analysis')

# ====== 9. DISCOUNT IMPACT ======
fig, axes = plt.subplots(2, 2, figsize=(16, 10))
fig.suptitle('Discount Impact Analysis', fontsize=16, fontweight='bold', y=1.01)

df['Disc_Band'] = pd.cut(df['Discount'], bins=[-0.01,0,0.1,0.2,0.3,0.5,1], labels=['None','1-10%','11-20%','21-30%','31-50%','50%+'])
disc_prof = df.groupby('Disc_Band', observed=True).agg({'Profit': 'sum', 'Sales': 'sum'}).reset_index()
disc_prof['Margin'] = disc_prof['Profit'] / disc_prof['Sales'] * 100
colors_d = [COLORS[0] if m>0 else COLORS[3] for m in disc_prof['Margin']]
axes[0,0].bar(disc_prof['Disc_Band'].astype(str), disc_prof['Margin'], color=colors_d)
axes[0,0].set_title('Profit Margin by Discount Band')
axes[0,0].axhline(y=0, color='black', linewidth=0.5)
axes[0,0].set_ylabel('Profit Margin (%)')

axes[0,1].scatter(df['Discount']*100, df['Profit'], alpha=0.15, s=8, color=COLORS[0])
z = np.polyfit(df['Discount']*100, df['Profit'], 1)
p = np.poly1d(z)
x_line = np.linspace(0, df['Discount'].max()*100, 100)
axes[0,1].plot(x_line, p(x_line), color=COLORS[3], linewidth=2, linestyle='--')
axes[0,1].set_title('Discount vs Profit (with trend line)')
axes[0,1].set_xlabel('Discount (%)'); axes[0,1].set_ylabel('Profit ($)')
axes[0,1].axhline(y=0, color='gray', linewidth=0.5, linestyle='--')

disc_cat = df.pivot_table(values='Profit', index='Disc_Band', columns='Category', aggfunc='sum', observed=True)
disc_cat.plot(kind='bar', ax=axes[1,0], color=COLORS[:3])
axes[1,0].set_title('Profit by Discount Band & Category')
axes[1,0].yaxis.set_major_formatter(mticker.FuncFormatter(currency_fmt))
axes[1,0].tick_params(axis='x', rotation=0)
axes[1,0].legend(fontsize=9)

disc_count = df['Disc_Band'].value_counts().sort_index()
axes[1,1].bar(disc_count.index.astype(str), disc_count.values, color=COLORS[5])
axes[1,1].set_title('Order Volume by Discount Band')
axes[1,1].set_ylabel('Number of Orders')

plt.tight_layout()
save(fig, '09_discount_impact')

# ====== 10. CORRELATION HEATMAP ======
fig, ax = plt.subplots(figsize=(10, 8))
corr_cols = ['Sales', 'Quantity', 'Discount', 'Profit', 'Delivery_Days', 'Profit_Margin']
corr = df[corr_cols].corr()
mask = np.triu(np.ones_like(corr, dtype=bool))
sns.heatmap(corr, mask=mask, annot=True, fmt='.2f', cmap='RdBu_r', center=0,
            square=True, linewidths=1, ax=ax, vmin=-1, vmax=1,
            cbar_kws={'shrink': 0.8})
ax.set_title('Correlation Matrix — Key Metrics', fontsize=14, fontweight='bold', pad=15)
save(fig, '10_correlation_heatmap')

print('All visualizations saved to images/ folder')
print(f'Generated 10 chart images')
