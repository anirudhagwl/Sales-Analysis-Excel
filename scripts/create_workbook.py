import pandas as pd
import numpy as np
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side, numbers
from openpyxl.utils import get_column_letter
from openpyxl.chart import BarChart, LineChart, PieChart, Reference, AreaChart
from openpyxl.chart.label import DataLabelList
from openpyxl.chart.series import DataPoint, SeriesLabel
from datetime import datetime

df = pd.read_csv('data/Sales_Orders.csv')
df['Order Date'] = pd.to_datetime(df['Order Date'], format='%d/%m/%Y')
df['Ship Date'] = pd.to_datetime(df['Ship Date'], format='%d/%m/%Y')
df['Year'] = df['Order Date'].dt.year
df['Month'] = df['Order Date'].dt.month
df['Month_Name'] = df['Order Date'].dt.strftime('%B')
df['Quarter'] = df['Order Date'].dt.quarter
df['Delivery_Days'] = (df['Ship Date'] - df['Order Date']).dt.days
df['Profit_Margin'] = df['Profit'] / df['Sales'] * 100

wb = Workbook()

DARK_BLUE = PatternFill('solid', fgColor='1F4E79')
LIGHT_BLUE = PatternFill('solid', fgColor='D6E4F0')
WHITE_FILL = PatternFill('solid', fgColor='FFFFFF')
LIGHT_GRAY = PatternFill('solid', fgColor='F2F2F2')
GREEN_FILL = PatternFill('solid', fgColor='E2EFDA')
RED_FILL = PatternFill('solid', fgColor='FCE4EC')
ACCENT_FILL = PatternFill('solid', fgColor='4472C4')
KPI_FILL = PatternFill('solid', fgColor='2E75B6')
HEADER_FONT = Font(name='Arial', bold=True, color='FFFFFF', size=11)
TITLE_FONT = Font(name='Arial', bold=True, color='1F4E79', size=14)
SUBTITLE_FONT = Font(name='Arial', bold=True, color='1F4E79', size=12)
DATA_FONT = Font(name='Arial', size=10)
CURRENCY_FMT = '$#,##0.00'
PCT_FMT = '0.0%'
NUM_FMT = '#,##0'
THIN_BORDER = Border(
    left=Side(style='thin', color='D9D9D9'),
    right=Side(style='thin', color='D9D9D9'),
    top=Side(style='thin', color='D9D9D9'),
    bottom=Side(style='thin', color='D9D9D9')
)

def style_header(ws, row, max_col):
    for col in range(1, max_col + 1):
        cell = ws.cell(row=row, column=col)
        cell.fill = DARK_BLUE
        cell.font = HEADER_FONT
        cell.alignment = Alignment(horizontal='center', vertical='center', wrap_text=True)
        cell.border = THIN_BORDER

def style_data_rows(ws, start_row, end_row, max_col):
    for r in range(start_row, end_row + 1):
        for c in range(1, max_col + 1):
            cell = ws.cell(row=r, column=c)
            cell.font = DATA_FONT
            cell.border = THIN_BORDER
            cell.alignment = Alignment(horizontal='center', vertical='center')
            if (r - start_row) % 2 == 0:
                cell.fill = WHITE_FILL
            else:
                cell.fill = LIGHT_GRAY

def add_title(ws, title, row=1, col=1):
    cell = ws.cell(row=row, column=col, value=title)
    cell.font = TITLE_FONT
    return row + 1

def auto_width(ws, max_col, min_width=12, max_width=30):
    for col in range(1, max_col + 1):
        ws.column_dimensions[get_column_letter(col)].width = min(max(min_width, 15), max_width)

# ========== 1. RAW DATA ==========
ws_raw = wb.active
ws_raw.title = 'Raw Data'
cols = ['Row ID','Order ID','Order Date','Ship Date','Ship Mode','Customer ID','Customer Name',
        'Segment','Country','City','State','Postal Code','Region','Category','Sub-Category',
        'Product Name','Sales','Quantity','Discount','Profit','Delivery_Days','Profit_Margin']
for c, header in enumerate(cols, 1):
    ws_raw.cell(row=1, column=c, value=header)
style_header(ws_raw, 1, len(cols))
for i, (_, row) in enumerate(df.iterrows(), 2):
    if i > 1001:
        break
    for c, col_name in enumerate(cols, 1):
        val = row[col_name]
        if pd.isna(val):
            val = ''
        elif isinstance(val, pd.Timestamp):
            val = val.strftime('%Y-%m-%d')
        ws_raw.cell(row=i, column=c, value=val)
style_data_rows(ws_raw, 2, min(1001, len(df)+1), len(cols))
for c in [17, 20]:
    for r in range(2, min(1002, len(df)+2)):
        ws_raw.cell(row=r, column=c).number_format = CURRENCY_FMT
for r in range(2, min(1002, len(df)+2)):
    ws_raw.cell(row=r, column=19).number_format = PCT_FMT
    ws_raw.cell(row=r, column=22).number_format = '0.0%'
ws_raw.freeze_panes = 'A2'
auto_width(ws_raw, len(cols))
ws_raw.cell(row=min(1003, len(df)+3), column=1, value=f'Showing first 1000 of {len(df)} records. Full data in CSV.').font = Font(name='Arial', italic=True, color='808080', size=9)

# ========== 2. SUMMARY STATISTICS ==========
ws_stats = wb.create_sheet('Summary Statistics')
add_title(ws_stats, 'Descriptive Statistics Summary', 1, 1)
stats_headers = ['Metric', 'Sales ($)', 'Profit ($)', 'Quantity', 'Discount', 'Delivery Days', 'Profit Margin (%)']
for c, h in enumerate(stats_headers, 1):
    ws_stats.cell(row=3, column=c, value=h)
style_header(ws_stats, 3, len(stats_headers))

stat_names = ['Count', 'Mean', 'Median', 'Std Dev', 'Min', 'Max', 'Q1 (25th)', 'Q3 (75th)', 'IQR', 'Skewness', 'Kurtosis', 'Sum', 'Variance', 'Coeff of Variation']
stat_cols = ['Sales', 'Profit', 'Quantity', 'Discount', 'Delivery_Days', 'Profit_Margin']
for r, stat in enumerate(stat_names, 4):
    ws_stats.cell(row=r, column=1, value=stat)
    for c, col_name in enumerate(stat_cols, 2):
        s = df[col_name].dropna()
        if stat == 'Count': val = len(s)
        elif stat == 'Mean': val = s.mean()
        elif stat == 'Median': val = s.median()
        elif stat == 'Std Dev': val = s.std()
        elif stat == 'Min': val = s.min()
        elif stat == 'Max': val = s.max()
        elif stat == 'Q1 (25th)': val = s.quantile(0.25)
        elif stat == 'Q3 (75th)': val = s.quantile(0.75)
        elif stat == 'IQR': val = s.quantile(0.75) - s.quantile(0.25)
        elif stat == 'Skewness': val = s.skew()
        elif stat == 'Kurtosis': val = s.kurtosis()
        elif stat == 'Sum': val = s.sum()
        elif stat == 'Variance': val = s.var()
        elif stat == 'Coeff of Variation': val = (s.std() / s.mean()) * 100 if s.mean() != 0 else 0
        ws_stats.cell(row=r, column=c, value=round(val, 4))

style_data_rows(ws_stats, 4, 4 + len(stat_names) - 1, len(stats_headers))
for r in range(4, 4 + len(stat_names)):
    ws_stats.cell(row=r, column=2).number_format = CURRENCY_FMT
    ws_stats.cell(row=r, column=3).number_format = CURRENCY_FMT
auto_width(ws_stats, len(stats_headers), 18)

add_title(ws_stats, 'Distribution by Segment', 20, 1)
seg_headers = ['Segment', 'Total Sales ($)', 'Total Profit ($)', 'Avg Order Value ($)', 'Order Count', 'Profit Margin (%)']
for c, h in enumerate(seg_headers, 1):
    ws_stats.cell(row=22, column=c, value=h)
style_header(ws_stats, 22, len(seg_headers))
seg_stats = df.groupby('Segment').agg({'Sales': ['sum', 'mean', 'count'], 'Profit': 'sum'}).reset_index()
seg_stats.columns = ['Segment', 'Total_Sales', 'Avg_Order', 'Count', 'Total_Profit']
seg_stats['Margin'] = seg_stats['Total_Profit'] / seg_stats['Total_Sales'] * 100
for r, (_, row) in enumerate(seg_stats.iterrows(), 23):
    ws_stats.cell(row=r, column=1, value=row['Segment'])
    ws_stats.cell(row=r, column=2, value=round(row['Total_Sales'], 2))
    ws_stats.cell(row=r, column=3, value=round(row['Total_Profit'], 2))
    ws_stats.cell(row=r, column=4, value=round(row['Avg_Order'], 2))
    ws_stats.cell(row=r, column=5, value=int(row['Count']))
    ws_stats.cell(row=r, column=6, value=round(row['Margin'], 2))
    ws_stats.cell(row=r, column=2).number_format = CURRENCY_FMT
    ws_stats.cell(row=r, column=3).number_format = CURRENCY_FMT
    ws_stats.cell(row=r, column=4).number_format = CURRENCY_FMT
style_data_rows(ws_stats, 23, 25, len(seg_headers))

# ========== 3. REVENUE ANALYSIS ==========
ws_rev = wb.create_sheet('Revenue Analysis')
add_title(ws_rev, 'Revenue Trend Analysis', 1, 1)

yearly = df.groupby('Year').agg({'Sales': 'sum', 'Profit': 'sum', 'Order ID': 'nunique', 'Quantity': 'sum'}).reset_index()
yearly.columns = ['Year', 'Revenue', 'Profit', 'Orders', 'Units']

rev_headers = ['Year', 'Total Revenue ($)', 'Total Profit ($)', 'Total Orders', 'Units Sold', 'YoY Revenue Growth (%)', 'YoY Profit Growth (%)']
for c, h in enumerate(rev_headers, 1):
    ws_rev.cell(row=3, column=c, value=h)
style_header(ws_rev, 3, len(rev_headers))

for r, (_, row) in enumerate(yearly.iterrows(), 4):
    ws_rev.cell(row=r, column=1, value=int(row['Year']))
    ws_rev.cell(row=r, column=2, value=round(row['Revenue'], 2))
    ws_rev.cell(row=r, column=3, value=round(row['Profit'], 2))
    ws_rev.cell(row=r, column=4, value=int(row['Orders']))
    ws_rev.cell(row=r, column=5, value=int(row['Units']))
    ws_rev.cell(row=r, column=2).number_format = CURRENCY_FMT
    ws_rev.cell(row=r, column=3).number_format = CURRENCY_FMT
    if r > 4:
        prev_rev = yearly.iloc[r-5]['Revenue']
        prev_prof = yearly.iloc[r-5]['Profit']
        ws_rev.cell(row=r, column=6, value=round((row['Revenue'] - prev_rev) / prev_rev * 100, 2))
        ws_rev.cell(row=r, column=7, value=round((row['Profit'] - prev_prof) / prev_prof * 100, 2))
    else:
        ws_rev.cell(row=r, column=6, value='N/A')
        ws_rev.cell(row=r, column=7, value='N/A')
style_data_rows(ws_rev, 4, 4 + len(yearly) - 1, len(rev_headers))

chart = BarChart()
chart.type = "col"
chart.title = "Yearly Revenue & Profit"
chart.y_axis.title = "Amount ($)"
chart.x_axis.title = "Year"
chart.style = 10
cats = Reference(ws_rev, min_col=1, min_row=4, max_row=4+len(yearly)-1)
rev_data = Reference(ws_rev, min_col=2, min_row=3, max_row=4+len(yearly)-1)
prof_data = Reference(ws_rev, min_col=3, min_row=3, max_row=4+len(yearly)-1)
chart.add_data(rev_data, titles_from_data=True)
chart.add_data(prof_data, titles_from_data=True)
chart.set_categories(cats)
chart.shape = 4
chart.width = 20
chart.height = 12
ws_rev.add_chart(chart, "A10")

# Monthly revenue
add_title(ws_rev, 'Monthly Revenue Breakdown', 28, 1)
monthly = df.groupby([df['Year'], df['Month']]).agg({'Sales': 'sum'}).reset_index()
months = ['Jan','Feb','Mar','Apr','May','Jun','Jul','Aug','Sep','Oct','Nov','Dec']
mon_headers = ['Year'] + months
for c, h in enumerate(mon_headers, 1):
    ws_rev.cell(row=30, column=c, value=h)
style_header(ws_rev, 30, len(mon_headers))

for yr_idx, year in enumerate(sorted(df['Year'].unique()), 31):
    ws_rev.cell(row=yr_idx, column=1, value=int(year))
    yr_data = monthly[monthly['Year'] == year]
    for _, row in yr_data.iterrows():
        ws_rev.cell(row=yr_idx, column=int(row['Month']) + 1, value=round(row['Sales'], 2))
        ws_rev.cell(row=yr_idx, column=int(row['Month']) + 1).number_format = CURRENCY_FMT
style_data_rows(ws_rev, 31, 31 + len(df['Year'].unique()) - 1, len(mon_headers))
auto_width(ws_rev, len(mon_headers))

line_chart = LineChart()
line_chart.title = "Monthly Revenue by Year"
line_chart.y_axis.title = "Revenue ($)"
line_chart.style = 10
line_chart.width = 24
line_chart.height = 14
cats = Reference(ws_rev, min_col=2, min_row=30, max_col=13, max_row=30)
for yr_idx, year in enumerate(sorted(df['Year'].unique())):
    data = Reference(ws_rev, min_col=2, max_col=13, min_row=31+yr_idx, max_row=31+yr_idx)
    line_chart.add_data(data, from_rows=True, titles_from_data=False)
    line_chart.series[yr_idx].tx = SeriesLabel(v=str(int(year)))
line_chart.set_categories(cats)
ws_rev.add_chart(line_chart, "A36")

# ========== 4. PROFITABILITY ANALYSIS ==========
ws_prof = wb.create_sheet('Profitability Analysis')
add_title(ws_prof, 'Profitability Analysis', 1, 1)

cat_prof = df.groupby('Category').agg({'Sales': 'sum', 'Profit': 'sum', 'Quantity': 'sum'}).reset_index()
cat_prof['Margin'] = cat_prof['Profit'] / cat_prof['Sales'] * 100

prof_h = ['Category', 'Total Sales ($)', 'Total Profit ($)', 'Profit Margin (%)', 'Units Sold']
for c, h in enumerate(prof_h, 1):
    ws_prof.cell(row=3, column=c, value=h)
style_header(ws_prof, 3, len(prof_h))
for r, (_, row) in enumerate(cat_prof.iterrows(), 4):
    ws_prof.cell(row=r, column=1, value=row['Category'])
    ws_prof.cell(row=r, column=2, value=round(row['Sales'], 2))
    ws_prof.cell(row=r, column=3, value=round(row['Profit'], 2))
    ws_prof.cell(row=r, column=4, value=round(row['Margin'], 2))
    ws_prof.cell(row=r, column=5, value=int(row['Quantity']))
    ws_prof.cell(row=r, column=2).number_format = CURRENCY_FMT
    ws_prof.cell(row=r, column=3).number_format = CURRENCY_FMT
style_data_rows(ws_prof, 4, 4 + len(cat_prof) - 1, len(prof_h))

pie = PieChart()
pie.title = "Sales Distribution by Category"
pie.style = 10
pie.width = 16
pie.height = 12
cats_ref = Reference(ws_prof, min_col=1, min_row=4, max_row=4+len(cat_prof)-1)
vals_ref = Reference(ws_prof, min_col=2, min_row=4, max_row=4+len(cat_prof)-1)
pie.add_data(vals_ref)
pie.set_categories(cats_ref)
pie.dataLabels = DataLabelList()
pie.dataLabels.showPercent = True
pie.dataLabels.showCatName = True
ws_prof.add_chart(pie, "G3")

# Sub-category profitability
add_title(ws_prof, 'Sub-Category Profitability', 20, 1)
subcat = df.groupby('Sub-Category').agg({'Sales': 'sum', 'Profit': 'sum'}).reset_index()
subcat['Margin'] = subcat['Profit'] / subcat['Sales'] * 100
subcat = subcat.sort_values('Profit', ascending=False)

subh = ['Sub-Category', 'Total Sales ($)', 'Total Profit ($)', 'Profit Margin (%)']
for c, h in enumerate(subh, 1):
    ws_prof.cell(row=22, column=c, value=h)
style_header(ws_prof, 22, len(subh))
for r, (_, row) in enumerate(subcat.iterrows(), 23):
    ws_prof.cell(row=r, column=1, value=row['Sub-Category'])
    ws_prof.cell(row=r, column=2, value=round(row['Sales'], 2))
    ws_prof.cell(row=r, column=3, value=round(row['Profit'], 2))
    ws_prof.cell(row=r, column=4, value=round(row['Margin'], 2))
    ws_prof.cell(row=r, column=2).number_format = CURRENCY_FMT
    ws_prof.cell(row=r, column=3).number_format = CURRENCY_FMT
    if row['Profit'] < 0:
        ws_prof.cell(row=r, column=3).font = Font(name='Arial', size=10, color='FF0000', bold=True)
style_data_rows(ws_prof, 23, 23 + len(subcat) - 1, len(subh))

bar = BarChart()
bar.type = "col"
bar.title = "Profit by Sub-Category"
bar.style = 10
bar.width = 22
bar.height = 14
cats_ref = Reference(ws_prof, min_col=1, min_row=23, max_row=23+len(subcat)-1)
vals_ref = Reference(ws_prof, min_col=3, min_row=22, max_row=23+len(subcat)-1)
bar.add_data(vals_ref, titles_from_data=True)
bar.set_categories(cats_ref)
ws_prof.add_chart(bar, "F22")
auto_width(ws_prof, 5)

# ========== 5. CUSTOMER SEGMENTATION (RFM) ==========
ws_cust = wb.create_sheet('Customer Segmentation')
add_title(ws_cust, 'RFM Customer Segmentation Analysis', 1, 1)

max_date = df['Order Date'].max()
rfm = df.groupby('Customer ID').agg({
    'Order Date': lambda x: (max_date - x.max()).days,
    'Order ID': 'nunique',
    'Sales': 'sum'
}).reset_index()
rfm.columns = ['Customer ID', 'Recency', 'Frequency', 'Monetary']
rfm['R_Score'] = pd.qcut(rfm['Recency'], 4, labels=[4,3,2,1]).astype(int)
rfm['F_Score'] = pd.qcut(rfm['Frequency'].rank(method='first'), 4, labels=[1,2,3,4]).astype(int)
rfm['M_Score'] = pd.qcut(rfm['Monetary'], 4, labels=[1,2,3,4]).astype(int)
rfm['RFM_Score'] = rfm['R_Score'] + rfm['F_Score'] + rfm['M_Score']

def rfm_segment(row):
    if row['RFM_Score'] >= 10: return 'Champions'
    elif row['RFM_Score'] >= 8: return 'Loyal Customers'
    elif row['RFM_Score'] >= 6: return 'Potential Loyalists'
    elif row['RFM_Score'] >= 4: return 'At Risk'
    else: return 'Lost'
rfm['Segment'] = rfm.apply(rfm_segment, axis=1)

cust_name_map = df.drop_duplicates('Customer ID').set_index('Customer ID')['Customer Name'].to_dict()
rfm['Customer Name'] = rfm['Customer ID'].map(cust_name_map)

rfm_h = ['Customer ID', 'Customer Name', 'Recency (days)', 'Frequency', 'Monetary ($)', 'R Score', 'F Score', 'M Score', 'RFM Score', 'Segment']
for c, h in enumerate(rfm_h, 1):
    ws_cust.cell(row=3, column=c, value=h)
style_header(ws_cust, 3, len(rfm_h))

rfm_sorted = rfm.sort_values('RFM_Score', ascending=False).head(50)
for r, (_, row) in enumerate(rfm_sorted.iterrows(), 4):
    ws_cust.cell(row=r, column=1, value=row['Customer ID'])
    ws_cust.cell(row=r, column=2, value=row['Customer Name'])
    ws_cust.cell(row=r, column=3, value=int(row['Recency']))
    ws_cust.cell(row=r, column=4, value=int(row['Frequency']))
    ws_cust.cell(row=r, column=5, value=round(row['Monetary'], 2))
    ws_cust.cell(row=r, column=6, value=int(row['R_Score']))
    ws_cust.cell(row=r, column=7, value=int(row['F_Score']))
    ws_cust.cell(row=r, column=8, value=int(row['M_Score']))
    ws_cust.cell(row=r, column=9, value=int(row['RFM_Score']))
    ws_cust.cell(row=r, column=10, value=row['Segment'])
    ws_cust.cell(row=r, column=5).number_format = CURRENCY_FMT
style_data_rows(ws_cust, 4, 4 + min(49, len(rfm_sorted)-1), len(rfm_h))

# Segment summary
add_title(ws_cust, 'RFM Segment Distribution', 57, 1)
seg_dist = rfm.groupby('Segment').agg({'Customer ID': 'count', 'Monetary': ['sum', 'mean']}).reset_index()
seg_dist.columns = ['Segment', 'Customer Count', 'Total Revenue', 'Avg Revenue']
seg_dist = seg_dist.sort_values('Total Revenue', ascending=False)

seg_h = ['Segment', 'Customer Count', '% of Customers', 'Total Revenue ($)', 'Avg Revenue per Customer ($)']
for c, h in enumerate(seg_h, 1):
    ws_cust.cell(row=59, column=c, value=h)
style_header(ws_cust, 59, len(seg_h))
total_custs = seg_dist['Customer Count'].sum()
for r, (_, row) in enumerate(seg_dist.iterrows(), 60):
    ws_cust.cell(row=r, column=1, value=row['Segment'])
    ws_cust.cell(row=r, column=2, value=int(row['Customer Count']))
    ws_cust.cell(row=r, column=3, value=round(row['Customer Count']/total_custs*100, 1))
    ws_cust.cell(row=r, column=4, value=round(row['Total Revenue'], 2))
    ws_cust.cell(row=r, column=5, value=round(row['Avg Revenue'], 2))
    ws_cust.cell(row=r, column=4).number_format = CURRENCY_FMT
    ws_cust.cell(row=r, column=5).number_format = CURRENCY_FMT
style_data_rows(ws_cust, 60, 60 + len(seg_dist) - 1, len(seg_h))
auto_width(ws_cust, len(rfm_h), 15)

pie2 = PieChart()
pie2.title = "Customer Segments"
pie2.style = 10
pie2.width = 16
pie2.height = 12
cats_ref = Reference(ws_cust, min_col=1, min_row=60, max_row=60+len(seg_dist)-1)
vals_ref = Reference(ws_cust, min_col=2, min_row=60, max_row=60+len(seg_dist)-1)
pie2.add_data(vals_ref)
pie2.set_categories(cats_ref)
pie2.dataLabels = DataLabelList()
pie2.dataLabels.showPercent = True
ws_cust.add_chart(pie2, "G59")

# ========== 6. PRODUCT PERFORMANCE ==========
ws_prod = wb.create_sheet('Product Performance')
add_title(ws_prod, 'Product Performance Analysis', 1, 1)

# Top 20 products
top_prods = df.groupby('Product Name').agg({'Sales': 'sum', 'Profit': 'sum', 'Quantity': 'sum'}).reset_index()
top_prods['Margin'] = top_prods['Profit'] / top_prods['Sales'] * 100
top_prods = top_prods.sort_values('Sales', ascending=False)

add_title(ws_prod, 'Top 20 Best-Selling Products', 3, 1)
prod_h = ['Rank', 'Product Name', 'Total Sales ($)', 'Total Profit ($)', 'Units Sold', 'Profit Margin (%)']
for c, h in enumerate(prod_h, 1):
    ws_prod.cell(row=5, column=c, value=h)
style_header(ws_prod, 5, len(prod_h))
for r, (_, row) in enumerate(top_prods.head(20).iterrows(), 6):
    ws_prod.cell(row=r, column=1, value=r-5)
    ws_prod.cell(row=r, column=2, value=row['Product Name'][:60])
    ws_prod.cell(row=r, column=3, value=round(row['Sales'], 2))
    ws_prod.cell(row=r, column=4, value=round(row['Profit'], 2))
    ws_prod.cell(row=r, column=5, value=int(row['Quantity']))
    ws_prod.cell(row=r, column=6, value=round(row['Margin'], 2))
    ws_prod.cell(row=r, column=3).number_format = CURRENCY_FMT
    ws_prod.cell(row=r, column=4).number_format = CURRENCY_FMT
style_data_rows(ws_prod, 6, 25, len(prod_h))

# Bottom 10 (loss-making)
add_title(ws_prod, 'Bottom 10 Loss-Making Products', 28, 1)
bottom_prods = top_prods.sort_values('Profit').head(10)
for c, h in enumerate(prod_h, 1):
    ws_prod.cell(row=30, column=c, value=h)
style_header(ws_prod, 30, len(prod_h))
for r, (_, row) in enumerate(bottom_prods.iterrows(), 31):
    ws_prod.cell(row=r, column=1, value=r-30)
    ws_prod.cell(row=r, column=2, value=row['Product Name'][:60])
    ws_prod.cell(row=r, column=3, value=round(row['Sales'], 2))
    ws_prod.cell(row=r, column=4, value=round(row['Profit'], 2))
    ws_prod.cell(row=r, column=5, value=int(row['Quantity']))
    ws_prod.cell(row=r, column=6, value=round(row['Margin'], 2))
    ws_prod.cell(row=r, column=3).number_format = CURRENCY_FMT
    ws_prod.cell(row=r, column=4).number_format = CURRENCY_FMT
    ws_prod.cell(row=r, column=4).font = Font(name='Arial', size=10, color='FF0000', bold=True)
style_data_rows(ws_prod, 31, 40, len(prod_h))

# ABC Analysis (Pareto)
add_title(ws_prod, 'ABC Analysis (Pareto 80/20 Rule)', 43, 1)
abc = top_prods.sort_values('Sales', ascending=False).reset_index(drop=True)
abc['Cumulative_Sales'] = abc['Sales'].cumsum()
abc['Cumulative_Pct'] = abc['Cumulative_Sales'] / abc['Sales'].sum() * 100
abc['Class'] = abc['Cumulative_Pct'].apply(lambda x: 'A' if x <= 80 else ('B' if x <= 95 else 'C'))

abc_summary = abc.groupby('Class').agg({'Product Name': 'count', 'Sales': 'sum'}).reset_index()
abc_summary.columns = ['Class', 'Product Count', 'Total Sales']
abc_summary['% of Products'] = abc_summary['Product Count'] / abc_summary['Product Count'].sum() * 100
abc_summary['% of Sales'] = abc_summary['Total Sales'] / abc_summary['Total Sales'].sum() * 100

abc_h = ['Class', 'Product Count', '% of Products', 'Total Sales ($)', '% of Sales']
for c, h in enumerate(abc_h, 1):
    ws_prod.cell(row=45, column=c, value=h)
style_header(ws_prod, 45, len(abc_h))
for r, (_, row) in enumerate(abc_summary.iterrows(), 46):
    ws_prod.cell(row=r, column=1, value=row['Class'])
    ws_prod.cell(row=r, column=2, value=int(row['Product Count']))
    ws_prod.cell(row=r, column=3, value=round(row['% of Products'], 1))
    ws_prod.cell(row=r, column=4, value=round(row['Total Sales'], 2))
    ws_prod.cell(row=r, column=5, value=round(row['% of Sales'], 1))
    ws_prod.cell(row=r, column=4).number_format = CURRENCY_FMT
style_data_rows(ws_prod, 46, 48, len(abc_h))

ws_prod.column_dimensions['B'].width = 50
auto_width(ws_prod, len(prod_h))
ws_prod.column_dimensions['B'].width = 50

# ========== 7. REGIONAL ANALYSIS ==========
ws_reg = wb.create_sheet('Regional Analysis')
add_title(ws_reg, 'Regional Sales & Profit Analysis', 1, 1)

region_data = df.groupby('Region').agg({'Sales': 'sum', 'Profit': 'sum', 'Order ID': 'nunique', 'Quantity': 'sum'}).reset_index()
region_data.columns = ['Region', 'Sales', 'Profit', 'Orders', 'Quantity']
region_data['Margin'] = region_data['Profit'] / region_data['Sales'] * 100
region_data = region_data.sort_values('Sales', ascending=False)

reg_h = ['Region', 'Total Sales ($)', 'Total Profit ($)', 'Profit Margin (%)', 'Total Orders', 'Units Sold', 'Avg Order Value ($)']
for c, h in enumerate(reg_h, 1):
    ws_reg.cell(row=3, column=c, value=h)
style_header(ws_reg, 3, len(reg_h))
for r, (_, row) in enumerate(region_data.iterrows(), 4):
    ws_reg.cell(row=r, column=1, value=row['Region'])
    ws_reg.cell(row=r, column=2, value=round(row['Sales'], 2))
    ws_reg.cell(row=r, column=3, value=round(row['Profit'], 2))
    ws_reg.cell(row=r, column=4, value=round(row['Margin'], 2))
    ws_reg.cell(row=r, column=5, value=int(row['Orders']))
    ws_reg.cell(row=r, column=6, value=int(row['Quantity']))
    ws_reg.cell(row=r, column=7, value=round(row['Sales']/row['Orders'], 2))
    ws_reg.cell(row=r, column=2).number_format = CURRENCY_FMT
    ws_reg.cell(row=r, column=3).number_format = CURRENCY_FMT
    ws_reg.cell(row=r, column=7).number_format = CURRENCY_FMT
style_data_rows(ws_reg, 4, 7, len(reg_h))

bar_reg = BarChart()
bar_reg.type = "col"
bar_reg.title = "Sales & Profit by Region"
bar_reg.style = 10
bar_reg.width = 18
bar_reg.height = 12
cats_ref = Reference(ws_reg, min_col=1, min_row=4, max_row=7)
s_ref = Reference(ws_reg, min_col=2, min_row=3, max_row=7)
p_ref = Reference(ws_reg, min_col=3, min_row=3, max_row=7)
bar_reg.add_data(s_ref, titles_from_data=True)
bar_reg.add_data(p_ref, titles_from_data=True)
bar_reg.set_categories(cats_ref)
ws_reg.add_chart(bar_reg, "A10")

# State-level analysis
add_title(ws_reg, 'Top 15 States by Revenue', 28, 1)
state_data = df.groupby('State').agg({'Sales': 'sum', 'Profit': 'sum', 'Order ID': 'nunique'}).reset_index()
state_data = state_data.sort_values('Sales', ascending=False).head(15)
state_data['Margin'] = state_data['Profit'] / state_data['Sales'] * 100

state_h = ['Rank', 'State', 'Total Sales ($)', 'Total Profit ($)', 'Profit Margin (%)', 'Orders']
for c, h in enumerate(state_h, 1):
    ws_reg.cell(row=30, column=c, value=h)
style_header(ws_reg, 30, len(state_h))
for r, (_, row) in enumerate(state_data.iterrows(), 31):
    ws_reg.cell(row=r, column=1, value=r-30)
    ws_reg.cell(row=r, column=2, value=row['State'])
    ws_reg.cell(row=r, column=3, value=round(row['Sales'], 2))
    ws_reg.cell(row=r, column=4, value=round(row['Profit'], 2))
    ws_reg.cell(row=r, column=5, value=round(row['Margin'], 2))
    ws_reg.cell(row=r, column=6, value=int(row['Order ID']))
    ws_reg.cell(row=r, column=3).number_format = CURRENCY_FMT
    ws_reg.cell(row=r, column=4).number_format = CURRENCY_FMT
    if row['Profit'] < 0:
        ws_reg.cell(row=r, column=4).font = Font(name='Arial', size=10, color='FF0000', bold=True)
style_data_rows(ws_reg, 31, 45, len(state_h))
auto_width(ws_reg, len(reg_h))

# ========== 8. TIME SERIES ANALYSIS ==========
ws_ts = wb.create_sheet('Time Series Analysis')
add_title(ws_ts, 'Time Series & Seasonality Analysis', 1, 1)

monthly_ts = df.groupby([df['Year'], df['Month']]).agg({'Sales': 'sum'}).reset_index()
monthly_ts.columns = ['Year', 'Month', 'Sales']
monthly_ts = monthly_ts.sort_values(['Year', 'Month'])
monthly_ts['Period'] = monthly_ts['Year'].astype(str) + '-' + monthly_ts['Month'].astype(str).str.zfill(2)

ts_h = ['Period', 'Revenue ($)', '3-Month MA ($)', '6-Month MA ($)', 'MoM Growth (%)']
for c, h in enumerate(ts_h, 1):
    ws_ts.cell(row=3, column=c, value=h)
style_header(ws_ts, 3, len(ts_h))

sales_list = monthly_ts['Sales'].tolist()
for r, (_, row) in enumerate(monthly_ts.iterrows(), 4):
    idx = r - 4
    ws_ts.cell(row=r, column=1, value=row['Period'])
    ws_ts.cell(row=r, column=2, value=round(row['Sales'], 2))
    ws_ts.cell(row=r, column=2).number_format = CURRENCY_FMT
    if idx >= 2:
        ma3 = np.mean(sales_list[max(0,idx-2):idx+1])
        ws_ts.cell(row=r, column=3, value=round(ma3, 2))
        ws_ts.cell(row=r, column=3).number_format = CURRENCY_FMT
    if idx >= 5:
        ma6 = np.mean(sales_list[max(0,idx-5):idx+1])
        ws_ts.cell(row=r, column=4, value=round(ma6, 2))
        ws_ts.cell(row=r, column=4).number_format = CURRENCY_FMT
    if idx > 0:
        mom = (sales_list[idx] - sales_list[idx-1]) / sales_list[idx-1] * 100
        ws_ts.cell(row=r, column=5, value=round(mom, 2))
end_row = 3 + len(monthly_ts)
style_data_rows(ws_ts, 4, end_row, len(ts_h))

line_ts = LineChart()
line_ts.title = "Revenue Trend with Moving Averages"
line_ts.y_axis.title = "Revenue ($)"
line_ts.style = 10
line_ts.width = 28
line_ts.height = 14
cats = Reference(ws_ts, min_col=1, min_row=4, max_row=end_row)
for col_idx in [2, 3, 4]:
    data = Reference(ws_ts, min_col=col_idx, min_row=3, max_row=end_row)
    line_ts.add_data(data, titles_from_data=True)
line_ts.set_categories(cats)
ws_ts.add_chart(line_ts, "A" + str(end_row + 3))

# Seasonality
season_row = end_row + 22
add_title(ws_ts, 'Seasonality Analysis (Avg Monthly Revenue)', season_row, 1)
season = df.groupby('Month').agg({'Sales': 'mean'}).reset_index()
seas_h = ['Month', 'Month Name', 'Avg Revenue ($)', 'Seasonality Index']
avg_monthly = season['Sales'].mean()
for c, h in enumerate(seas_h, 1):
    ws_ts.cell(row=season_row+2, column=c, value=h)
style_header(ws_ts, season_row+2, len(seas_h))
for r, (_, row) in enumerate(season.iterrows(), season_row+3):
    ws_ts.cell(row=r, column=1, value=int(row['Month']))
    ws_ts.cell(row=r, column=2, value=months[int(row['Month'])-1])
    ws_ts.cell(row=r, column=3, value=round(row['Sales'], 2))
    ws_ts.cell(row=r, column=3).number_format = CURRENCY_FMT
    ws_ts.cell(row=r, column=4, value=round(row['Sales']/avg_monthly, 3))
style_data_rows(ws_ts, season_row+3, season_row+14, len(seas_h))
auto_width(ws_ts, len(ts_h))

# ========== 9. SHIPPING ANALYSIS ==========
ws_ship = wb.create_sheet('Shipping Analysis')
add_title(ws_ship, 'Shipping & Delivery Analysis', 1, 1)

ship_mode = df.groupby('Ship Mode').agg({
    'Sales': 'sum', 'Profit': 'sum', 'Order ID': 'nunique',
    'Delivery_Days': 'mean', 'Quantity': 'sum'
}).reset_index()
ship_mode.columns = ['Ship Mode', 'Sales', 'Profit', 'Orders', 'Avg Delivery Days', 'Quantity']
ship_mode = ship_mode.sort_values('Sales', ascending=False)

ship_h = ['Ship Mode', 'Total Sales ($)', 'Total Profit ($)', 'Orders', '% of Orders', 'Avg Delivery Days', 'Units Shipped']
for c, h in enumerate(ship_h, 1):
    ws_ship.cell(row=3, column=c, value=h)
style_header(ws_ship, 3, len(ship_h))
total_orders = ship_mode['Orders'].sum()
for r, (_, row) in enumerate(ship_mode.iterrows(), 4):
    ws_ship.cell(row=r, column=1, value=row['Ship Mode'])
    ws_ship.cell(row=r, column=2, value=round(row['Sales'], 2))
    ws_ship.cell(row=r, column=3, value=round(row['Profit'], 2))
    ws_ship.cell(row=r, column=4, value=int(row['Orders']))
    ws_ship.cell(row=r, column=5, value=round(row['Orders']/total_orders*100, 1))
    ws_ship.cell(row=r, column=6, value=round(row['Avg Delivery Days'], 1))
    ws_ship.cell(row=r, column=7, value=int(row['Quantity']))
    ws_ship.cell(row=r, column=2).number_format = CURRENCY_FMT
    ws_ship.cell(row=r, column=3).number_format = CURRENCY_FMT
style_data_rows(ws_ship, 4, 7, len(ship_h))

pie_ship = PieChart()
pie_ship.title = "Orders by Shipping Mode"
pie_ship.style = 10
pie_ship.width = 16
pie_ship.height = 12
cats_ref = Reference(ws_ship, min_col=1, min_row=4, max_row=7)
vals_ref = Reference(ws_ship, min_col=4, min_row=4, max_row=7)
pie_ship.add_data(vals_ref)
pie_ship.set_categories(cats_ref)
pie_ship.dataLabels = DataLabelList()
pie_ship.dataLabels.showPercent = True
ws_ship.add_chart(pie_ship, "A10")

# Delivery time distribution
add_title(ws_ship, 'Delivery Time Distribution', 28, 1)
delivery_bins = pd.cut(df['Delivery_Days'], bins=[0,2,4,6,8,100], labels=['0-2 days','3-4 days','5-6 days','7-8 days','8+ days'])
del_dist = delivery_bins.value_counts().sort_index()
del_h = ['Delivery Window', 'Order Count', '% of Total']
for c, h in enumerate(del_h, 1):
    ws_ship.cell(row=30, column=c, value=h)
style_header(ws_ship, 30, len(del_h))
for r, (label, count) in enumerate(del_dist.items(), 31):
    ws_ship.cell(row=r, column=1, value=str(label))
    ws_ship.cell(row=r, column=2, value=int(count))
    ws_ship.cell(row=r, column=3, value=round(count/len(df)*100, 1))
style_data_rows(ws_ship, 31, 35, len(del_h))
auto_width(ws_ship, len(ship_h))

# ========== 10. DISCOUNT IMPACT ==========
ws_disc = wb.create_sheet('Discount Impact')
add_title(ws_disc, 'Discount Impact on Profitability', 1, 1)

df['Discount_Band'] = pd.cut(df['Discount'], bins=[-0.01, 0, 0.1, 0.2, 0.3, 0.5, 1.0],
                              labels=['No Discount', '1-10%', '11-20%', '21-30%', '31-50%', '50%+'])
disc_analysis = df.groupby('Discount_Band', observed=True).agg({
    'Sales': ['sum', 'mean', 'count'],
    'Profit': ['sum', 'mean'],
    'Quantity': 'sum'
}).reset_index()
disc_analysis.columns = ['Discount Band', 'Total Sales', 'Avg Sale', 'Order Count', 'Total Profit', 'Avg Profit', 'Quantity']
disc_analysis['Profit Margin'] = disc_analysis['Total Profit'] / disc_analysis['Total Sales'] * 100

disc_h = ['Discount Band', 'Order Count', '% of Orders', 'Total Sales ($)', 'Avg Sale ($)', 'Total Profit ($)', 'Avg Profit ($)', 'Profit Margin (%)']
for c, h in enumerate(disc_h, 1):
    ws_disc.cell(row=3, column=c, value=h)
style_header(ws_disc, 3, len(disc_h))
total_count = disc_analysis['Order Count'].sum()
for r, (_, row) in enumerate(disc_analysis.iterrows(), 4):
    ws_disc.cell(row=r, column=1, value=str(row['Discount Band']))
    ws_disc.cell(row=r, column=2, value=int(row['Order Count']))
    ws_disc.cell(row=r, column=3, value=round(row['Order Count']/total_count*100, 1))
    ws_disc.cell(row=r, column=4, value=round(row['Total Sales'], 2))
    ws_disc.cell(row=r, column=5, value=round(row['Avg Sale'], 2))
    ws_disc.cell(row=r, column=6, value=round(row['Total Profit'], 2))
    ws_disc.cell(row=r, column=7, value=round(row['Avg Profit'], 2))
    ws_disc.cell(row=r, column=8, value=round(row['Profit Margin'], 2))
    for col in [4,5,6,7]:
        ws_disc.cell(row=r, column=col).number_format = CURRENCY_FMT
    if row['Total Profit'] < 0:
        ws_disc.cell(row=r, column=6).font = Font(name='Arial', size=10, color='FF0000', bold=True)
end_disc = 3 + len(disc_analysis)
style_data_rows(ws_disc, 4, end_disc, len(disc_h))

bar_disc = BarChart()
bar_disc.type = "col"
bar_disc.title = "Profit Margin by Discount Band"
bar_disc.style = 10
bar_disc.width = 20
bar_disc.height = 12
cats_ref = Reference(ws_disc, min_col=1, min_row=4, max_row=end_disc)
vals_ref = Reference(ws_disc, min_col=8, min_row=3, max_row=end_disc)
bar_disc.add_data(vals_ref, titles_from_data=True)
bar_disc.set_categories(cats_ref)
ws_disc.add_chart(bar_disc, "A12")

# Discount by category
add_title(ws_disc, 'Discount Analysis by Category', 30, 1)
cat_disc = df.groupby(['Category', 'Discount_Band'], observed=True).agg({'Sales': 'sum', 'Profit': 'sum'}).reset_index()
cat_disc['Margin'] = cat_disc['Profit'] / cat_disc['Sales'] * 100
catd_h = ['Category', 'Discount Band', 'Total Sales ($)', 'Total Profit ($)', 'Profit Margin (%)']
for c, h in enumerate(catd_h, 1):
    ws_disc.cell(row=32, column=c, value=h)
style_header(ws_disc, 32, len(catd_h))
for r, (_, row) in enumerate(cat_disc.iterrows(), 33):
    ws_disc.cell(row=r, column=1, value=row['Category'])
    ws_disc.cell(row=r, column=2, value=str(row['Discount_Band']))
    ws_disc.cell(row=r, column=3, value=round(row['Sales'], 2))
    ws_disc.cell(row=r, column=4, value=round(row['Profit'], 2))
    ws_disc.cell(row=r, column=5, value=round(row['Margin'], 2))
    ws_disc.cell(row=r, column=3).number_format = CURRENCY_FMT
    ws_disc.cell(row=r, column=4).number_format = CURRENCY_FMT
style_data_rows(ws_disc, 33, 33 + len(cat_disc) - 1, len(catd_h))
auto_width(ws_disc, len(disc_h))

# ========== 11. KPI DASHBOARD ==========
ws_kpi = wb.create_sheet('KPI Dashboard')
add_title(ws_kpi, 'Executive KPI Dashboard', 1, 1)
ws_kpi.cell(row=2, column=1, value=f'Data Period: {df["Order Date"].min().strftime("%b %Y")} - {df["Order Date"].max().strftime("%b %Y")}').font = Font(name='Arial', size=10, italic=True, color='666666')

kpi_data = [
    ('Total Revenue', df['Sales'].sum(), CURRENCY_FMT),
    ('Total Profit', df['Profit'].sum(), CURRENCY_FMT),
    ('Overall Profit Margin', df['Profit'].sum()/df['Sales'].sum()*100, '0.0"%"'),
    ('Total Orders', df['Order ID'].nunique(), NUM_FMT),
    ('Total Customers', df['Customer ID'].nunique(), NUM_FMT),
    ('Total Products', df['Product Name'].nunique(), NUM_FMT),
    ('Avg Order Value', df['Sales'].sum()/df['Order ID'].nunique(), CURRENCY_FMT),
    ('Avg Profit per Order', df['Profit'].sum()/df['Order ID'].nunique(), CURRENCY_FMT),
    ('Avg Discount', df['Discount'].mean()*100, '0.0"%"'),
    ('Avg Delivery Days', df['Delivery_Days'].mean(), '0.0'),
    ('Total Units Sold', df['Quantity'].sum(), NUM_FMT),
    ('Avg Units per Order', df['Quantity'].sum()/df['Order ID'].nunique(), '0.0'),
]

for i, (label, value, fmt) in enumerate(kpi_data):
    row_offset = 4 + (i // 3) * 3
    col_offset = 1 + (i % 3) * 3
    cell_label = ws_kpi.cell(row=row_offset, column=col_offset, value=label)
    cell_label.font = Font(name='Arial', size=9, color='FFFFFF', bold=True)
    cell_label.fill = KPI_FILL
    cell_label.alignment = Alignment(horizontal='center')
    ws_kpi.merge_cells(start_row=row_offset, start_column=col_offset, end_row=row_offset, end_column=col_offset+1)
    cell_val = ws_kpi.cell(row=row_offset+1, column=col_offset, value=round(value, 2))
    cell_val.font = Font(name='Arial', size=16, bold=True, color='1F4E79')
    cell_val.alignment = Alignment(horizontal='center')
    cell_val.number_format = fmt
    ws_kpi.merge_cells(start_row=row_offset+1, start_column=col_offset, end_row=row_offset+1, end_column=col_offset+1)

# Category breakdown
kpi_cat_row = 20
add_title(ws_kpi, 'Category Performance Summary', kpi_cat_row, 1)
cat_h = ['Category', 'Revenue ($)', 'Profit ($)', 'Margin (%)', 'Orders', '% of Revenue']
for c, h in enumerate(cat_h, 1):
    ws_kpi.cell(row=kpi_cat_row+2, column=c, value=h)
style_header(ws_kpi, kpi_cat_row+2, len(cat_h))
total_rev = df['Sales'].sum()
for r, (_, row) in enumerate(cat_prof.sort_values('Sales', ascending=False).iterrows(), kpi_cat_row+3):
    ws_kpi.cell(row=r, column=1, value=row['Category'])
    ws_kpi.cell(row=r, column=2, value=round(row['Sales'], 2))
    ws_kpi.cell(row=r, column=3, value=round(row['Profit'], 2))
    ws_kpi.cell(row=r, column=4, value=round(row['Margin'], 2))
    ws_kpi.cell(row=r, column=5, value=int(row['Quantity']))
    ws_kpi.cell(row=r, column=6, value=round(row['Sales']/total_rev*100, 1))
    ws_kpi.cell(row=r, column=2).number_format = CURRENCY_FMT
    ws_kpi.cell(row=r, column=3).number_format = CURRENCY_FMT
style_data_rows(ws_kpi, kpi_cat_row+3, kpi_cat_row+5, len(cat_h))

# Regional KPIs
reg_kpi_row = kpi_cat_row + 8
add_title(ws_kpi, 'Regional Performance Summary', reg_kpi_row, 1)
reg_h2 = ['Region', 'Revenue ($)', 'Profit ($)', 'Margin (%)', 'Orders', '% of Revenue']
for c, h in enumerate(reg_h2, 1):
    ws_kpi.cell(row=reg_kpi_row+2, column=c, value=h)
style_header(ws_kpi, reg_kpi_row+2, len(reg_h2))
for r, (_, row) in enumerate(region_data.iterrows(), reg_kpi_row+3):
    ws_kpi.cell(row=r, column=1, value=row['Region'])
    ws_kpi.cell(row=r, column=2, value=round(row['Sales'], 2))
    ws_kpi.cell(row=r, column=3, value=round(row['Profit'], 2))
    ws_kpi.cell(row=r, column=4, value=round(row['Margin'], 2))
    ws_kpi.cell(row=r, column=5, value=int(row['Orders']))
    ws_kpi.cell(row=r, column=6, value=round(row['Sales']/total_rev*100, 1))
    ws_kpi.cell(row=r, column=2).number_format = CURRENCY_FMT
    ws_kpi.cell(row=r, column=3).number_format = CURRENCY_FMT
style_data_rows(ws_kpi, reg_kpi_row+3, reg_kpi_row+6, len(reg_h2))

for col in range(1, 10):
    ws_kpi.column_dimensions[get_column_letter(col)].width = 18

# ========== 12. COHORT ANALYSIS ==========
ws_cohort = wb.create_sheet('Cohort Analysis')
add_title(ws_cohort, 'Customer Cohort Retention Analysis', 1, 1)

df['Order_Month'] = df['Order Date'].dt.to_period('M')
first_purchase = df.groupby('Customer ID')['Order_Month'].min().reset_index()
first_purchase.columns = ['Customer ID', 'Cohort']
df_cohort = df.merge(first_purchase, on='Customer ID')
df_cohort['Cohort_Index'] = (df_cohort['Order_Month'] - df_cohort['Cohort']).apply(lambda x: x.n)

cohort_data = df_cohort.groupby(['Cohort', 'Cohort_Index'])['Customer ID'].nunique().reset_index()
cohort_data.columns = ['Cohort', 'Cohort_Index', 'Customers']
cohort_pivot = cohort_data.pivot(index='Cohort', columns='Cohort_Index', values='Customers').fillna(0)
cohort_sizes = cohort_pivot[0]
cohort_pct = cohort_pivot.div(cohort_sizes, axis=0) * 100

# Write quarterly cohort summary
quarterly_cohorts = {}
for cohort in cohort_pct.index:
    q = f"{cohort.year} Q{(cohort.month-1)//3+1}"
    if q not in quarterly_cohorts:
        quarterly_cohorts[q] = []
    quarterly_cohorts[q].append(cohort)

add_title(ws_cohort, 'Quarterly Cohort Retention (%)', 3, 1)
q_periods = ['Month 0', 'Month 3', 'Month 6', 'Month 9', 'Month 12', 'Month 18', 'Month 24']
coh_h = ['Cohort Quarter', 'New Customers'] + q_periods
for c, h in enumerate(coh_h, 1):
    ws_cohort.cell(row=5, column=c, value=h)
style_header(ws_cohort, 5, len(coh_h))

row_idx = 6
for q_name, cohorts in sorted(quarterly_cohorts.items()):
    total_new = sum(cohort_sizes.get(c, 0) for c in cohorts)
    ws_cohort.cell(row=row_idx, column=1, value=q_name)
    ws_cohort.cell(row=row_idx, column=2, value=int(total_new))
    for ci, month_idx in enumerate([0, 3, 6, 9, 12, 18, 24]):
        vals = []
        for c in cohorts:
            if month_idx in cohort_pct.columns:
                v = cohort_pct.loc[c, month_idx] if c in cohort_pct.index else 0
                if v > 0:
                    vals.append(v)
        if vals:
            ws_cohort.cell(row=row_idx, column=3+ci, value=round(np.mean(vals), 1))
    row_idx += 1
style_data_rows(ws_cohort, 6, row_idx-1, len(coh_h))
auto_width(ws_cohort, len(coh_h))

# Cohort revenue
add_title(ws_cohort, 'Revenue by Customer Cohort Quarter', row_idx + 2, 1)
rev_cohort = df_cohort.groupby(['Cohort', 'Cohort_Index'])['Sales'].sum().reset_index()
rev_pivot = rev_cohort.pivot(index='Cohort', columns='Cohort_Index', values='Sales').fillna(0)

rev_h = ['Cohort Quarter', 'Total Revenue ($)', 'Avg Revenue per Customer ($)', 'Customer Count']
for c, h in enumerate(rev_h, 1):
    ws_cohort.cell(row=row_idx+4, column=c, value=h)
style_header(ws_cohort, row_idx+4, len(rev_h))
rev_row = row_idx + 5
for q_name, cohorts in sorted(quarterly_cohorts.items()):
    total_rev_q = sum(rev_pivot.loc[c].sum() if c in rev_pivot.index else 0 for c in cohorts)
    total_cust = sum(cohort_sizes.get(c, 0) for c in cohorts)
    ws_cohort.cell(row=rev_row, column=1, value=q_name)
    ws_cohort.cell(row=rev_row, column=2, value=round(total_rev_q, 2))
    ws_cohort.cell(row=rev_row, column=3, value=round(total_rev_q/max(total_cust,1), 2))
    ws_cohort.cell(row=rev_row, column=4, value=int(total_cust))
    ws_cohort.cell(row=rev_row, column=2).number_format = CURRENCY_FMT
    ws_cohort.cell(row=rev_row, column=3).number_format = CURRENCY_FMT
    rev_row += 1
style_data_rows(ws_cohort, row_idx+5, rev_row-1, len(rev_h))

# Save
output_path = 'output/Sales_Analysis_Report.xlsx'
wb.save(output_path)
print(f'Workbook saved to {output_path}')
print(f'Total sheets: {len(wb.sheetnames)}')
print(f'Sheets: {", ".join(wb.sheetnames)}')
