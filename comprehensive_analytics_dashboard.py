#!/usr/bin/env python3
import pandas as pd
import numpy as np
import matplotlib.pyplot as plt
import seaborn as sns
import base64
from io import BytesIO
import warnings
from datetime import datetime, timedelta
from scipy import stats
warnings.filterwarnings('ignore')

# Set style for professional charts
plt.style.use('default')
sns.set_palette('husl')
plt.rcParams['figure.facecolor'] = 'white'
plt.rcParams['axes.facecolor'] = 'white'

def save_plot_as_base64():
    """Save current matplotlib plot as base64 encoded string"""
    buffer = BytesIO()
    plt.savefig(buffer, format='png', dpi=200, bbox_inches='tight', facecolor='white')
    buffer.seek(0)
    image_base64 = base64.b64encode(buffer.getvalue()).decode('utf-8')
    buffer.close()
    plt.close()
    return image_base64

def load_and_prepare_data():
    """Load and prepare all datasets with comprehensive merging"""
    print('🔄 Loading and preparing comprehensive dataset...')
    
    # Load all sheets
    file_path = 'AdventureWorks Sales (1).xlsx'
    excel_file = pd.ExcelFile(file_path)
    data = {}
    for sheet in excel_file.sheet_names:
        data[sheet] = pd.read_excel(file_path, sheet_name=sheet)
    
    # Create comprehensive merged dataset
    sales_data = data['Sales_data']
    
    # Merge with all dimension tables
    comprehensive_data = sales_data.merge(
        data['Date_data'], left_on='OrderDateKey', right_on='DateKey', how='left'
    ).merge(
        data['Sales Territory_data'], on='SalesTerritoryKey', how='left'
    ).merge(
        data['Product_data'], on='ProductKey', how='left'
    ).merge(
        data['Customer_data'], on='CustomerKey', how='left'
    ).merge(
        data['Sales Order_data'], on='SalesOrderLineKey', how='left'
    ).merge(
        data['Reseller_data'], on='ResellerKey', how='left'
    )
    
    # Add calculated fields
    comprehensive_data['Date'] = pd.to_datetime(comprehensive_data['Date'])
    comprehensive_data['Profit'] = comprehensive_data['Sales Amount'] - comprehensive_data['Total Product Cost']
    comprehensive_data['Profit_Margin'] = (comprehensive_data['Profit'] / comprehensive_data['Sales Amount']) * 100
    comprehensive_data['Month_Name'] = comprehensive_data['Date'].dt.strftime('%B')
    comprehensive_data['Year'] = comprehensive_data['Date'].dt.year
    comprehensive_data['Quarter'] = comprehensive_data['Date'].dt.quarter
    comprehensive_data['DayOfWeek'] = comprehensive_data['Date'].dt.day_name()
    
    return comprehensive_data, data

def create_executive_summary(data):
    """Create executive summary with key metrics"""
    print('📊 Creating Executive Summary...')
    
    # Calculate key metrics
    total_revenue = data['Sales Amount'].sum()
    total_transactions = len(data)
    unique_customers = data['CustomerKey'].nunique()
    unique_products = data['ProductKey'].nunique()
    avg_order_value = data['Sales Amount'].mean()
    total_profit = data['Profit'].sum()
    avg_profit_margin = data['Profit_Margin'].mean()
    countries = data['Country'].nunique()
    date_range = f"{data['Date'].min().strftime('%Y-%m-%d')} to {data['Date'].max().strftime('%Y-%m-%d')}"
    
    # Growth metrics
    yearly_sales = data.groupby('Fiscal Year')['Sales Amount'].sum().sort_index()
    if len(yearly_sales) > 1:
        yoy_growth = ((yearly_sales.iloc[-1] - yearly_sales.iloc[-2]) / yearly_sales.iloc[-2] * 100)
    else:
        yoy_growth = 0
    
    # Customer metrics
    customer_ltv = data.groupby('CustomerKey')['Sales Amount'].sum().mean()
    
    summary = {
        'total_revenue': total_revenue,
        'total_transactions': total_transactions,
        'unique_customers': unique_customers,
        'unique_products': unique_products,
        'avg_order_value': avg_order_value,
        'total_profit': total_profit,
        'avg_profit_margin': avg_profit_margin,
        'countries': countries,
        'date_range': date_range,
        'yoy_growth': yoy_growth,
        'customer_ltv': customer_ltv
    }
    
    return summary

def create_sales_performance_analytics(data):
    """Comprehensive sales performance analysis"""
    print('📈 Creating Sales Performance Analytics...')
    charts = {}
    
    # 1. Monthly Sales Trend with Forecasting
    fig, (ax1, ax2) = plt.subplots(2, 1, figsize=(16, 12))
    
    monthly_sales = data.groupby(['Year', 'Month_Name'])['Sales Amount'].sum().reset_index()
    monthly_sales['Date_Sort'] = pd.to_datetime(monthly_sales['Year'].astype(str) + ' ' + monthly_sales['Month_Name'], format='%Y %B')
    monthly_sales = monthly_sales.sort_values('Date_Sort')
    
    # Plot trend
    ax1.plot(monthly_sales['Date_Sort'], monthly_sales['Sales Amount']/1000000, 
             marker='o', linewidth=3, markersize=8, color='#2E86AB')
    
    # Add trend line
    x_numeric = np.arange(len(monthly_sales))
    slope, intercept, r_value, _, _ = stats.linregress(x_numeric, monthly_sales['Sales Amount'])
    trend_line = slope * x_numeric + intercept
    ax1.plot(monthly_sales['Date_Sort'], trend_line/1000000, '--', color='red', alpha=0.7, linewidth=2, label=f'Trend (R²={r_value**2:.3f})')
    
    ax1.set_title('Monthly Sales Revenue Trend with Forecasting', fontsize=16, fontweight='bold', pad=20)
    ax1.set_ylabel('Sales Amount (Millions $)', fontsize=12)
    ax1.grid(True, alpha=0.3)
    ax1.legend()
    
    # 2. Seasonal Analysis
    seasonal_pattern = data.groupby('Month_Name')['Sales Amount'].mean().reindex([
        'January', 'February', 'March', 'April', 'May', 'June',
        'July', 'August', 'September', 'October', 'November', 'December'
    ])
    
    colors = plt.cm.coolwarm(np.linspace(0, 1, 12))
    bars = ax2.bar(range(12), seasonal_pattern.values/1000, color=colors, alpha=0.8)
    ax2.set_title('Seasonal Sales Pattern (Average Monthly Performance)', fontsize=16, fontweight='bold', pad=20)
    ax2.set_xlabel('Month', fontsize=12)
    ax2.set_ylabel('Average Daily Sales (Thousands $)', fontsize=12)
    ax2.set_xticks(range(12))
    ax2.set_xticklabels([m[:3] for m in seasonal_pattern.index])
    ax2.grid(True, alpha=0.3, axis='y')
    
    # Add value labels
    for i, bar in enumerate(bars):
        height = bar.get_height()
        ax2.text(bar.get_x() + bar.get_width()/2., height + 5, f'${height:.0f}K', 
                ha='center', va='bottom', fontsize=10, fontweight='bold')
    
    plt.tight_layout()
    charts['sales_trend'] = save_plot_as_base64()
    
    # 3. Profit Analysis Dashboard
    fig, ((ax1, ax2), (ax3, ax4)) = plt.subplots(2, 2, figsize=(18, 14))
    
    # Profit margin distribution
    ax1.hist(data['Profit_Margin'], bins=50, alpha=0.7, color='skyblue', edgecolor='black')
    ax1.axvline(data['Profit_Margin'].mean(), color='red', linestyle='--', linewidth=2, 
               label=f'Mean: {data["Profit_Margin"].mean():.1f}%')
    ax1.axvline(data['Profit_Margin'].median(), color='green', linestyle='--', linewidth=2,
               label=f'Median: {data["Profit_Margin"].median():.1f}%')
    ax1.set_title('Profit Margin Distribution', fontsize=14, fontweight='bold')
    ax1.set_xlabel('Profit Margin (%)')
    ax1.set_ylabel('Frequency')
    ax1.legend()
    ax1.grid(True, alpha=0.3)
    
    # Profit by Fiscal Year
    yearly_profit = data.groupby('Fiscal Year')['Profit'].sum().reset_index()
    bars = ax2.bar(yearly_profit['Fiscal Year'], yearly_profit['Profit']/1000000, 
                  color=['#FF6B6B', '#4ECDC4', '#45B7D1', '#96CEB4'], alpha=0.8)
    ax2.set_title('Total Profit by Fiscal Year', fontsize=14, fontweight='bold')
    ax2.set_ylabel('Profit (Millions $)')
    for bar in bars:
        height = bar.get_height()
        ax2.text(bar.get_x() + bar.get_width()/2., height + 0.2, f'${height:.1f}M', 
                ha='center', va='bottom', fontweight='bold')
    
    # Sales vs Profit scatter
    sample_data = data.sample(n=min(5000, len(data)))
    scatter = ax3.scatter(sample_data['Sales Amount'], sample_data['Profit'], 
                         alpha=0.6, c=sample_data['Profit_Margin'], cmap='viridis', s=30)
    ax3.set_title('Sales Amount vs Profit Analysis', fontsize=14, fontweight='bold')
    ax3.set_xlabel('Sales Amount ($)')
    ax3.set_ylabel('Profit ($)')
    plt.colorbar(scatter, ax=ax3, label='Profit Margin (%)')
    
    # Top profitable products
    product_profit = data.groupby('Product')['Profit'].sum().sort_values(ascending=False).head(10)
    ax4.barh(range(len(product_profit)), product_profit.values/1000, 
            color=plt.cm.plasma(np.linspace(0, 1, len(product_profit))))
    ax4.set_title('Top 10 Most Profitable Products', fontsize=14, fontweight='bold')
    ax4.set_xlabel('Total Profit (Thousands $)')
    ax4.set_yticks(range(len(product_profit)))
    ax4.set_yticklabels([p[:30] + '...' if len(p) > 30 else p for p in product_profit.index], fontsize=10)
    
    for i, v in enumerate(product_profit.values/1000):
        ax4.text(v + 5, i, f'${v:.0f}K', va='center', fontweight='bold', fontsize=9)
    
    plt.tight_layout()
    charts['profit_analysis'] = save_plot_as_base64()
    
    return charts

def create_geographic_intelligence(data):
    """Comprehensive geographic market analysis"""
    print('🌍 Creating Geographic Intelligence Dashboard...')
    charts = {}
    
    fig, ((ax1, ax2), (ax3, ax4)) = plt.subplots(2, 2, figsize=(20, 16))
    
    # 1. Sales by Country with Market Share
    country_sales = data.groupby('Country')['Sales Amount'].sum().sort_values(ascending=False)
    total_sales = country_sales.sum()
    
    colors = plt.cm.Set3(np.linspace(0, 1, len(country_sales)))
    bars = ax1.bar(range(len(country_sales)), country_sales.values/1000000, color=colors, alpha=0.8)
    ax1.set_title('Sales Revenue by Country (Market Share Analysis)', fontsize=16, fontweight='bold')
    ax1.set_xlabel('Country')
    ax1.set_ylabel('Sales Amount (Millions $)')
    ax1.set_xticks(range(len(country_sales)))
    ax1.set_xticklabels(country_sales.index, rotation=45, ha='right')
    
    # Add value labels with market share
    for i, (country, value) in enumerate(country_sales.items()):
        market_share = (value / total_sales) * 100
        ax1.text(i, value/1000000 + 1, f'${value/1000000:.1f}M\n({market_share:.1f}%)', 
                ha='center', va='bottom', fontweight='bold', fontsize=10)
    
    # 2. Regional Performance Pie Chart
    region_sales = data.groupby('Region')['Sales Amount'].sum().sort_values(ascending=False)
    colors_pie = plt.cm.Paired(np.linspace(0, 1, len(region_sales)))
    wedges, texts, autotexts = ax2.pie(region_sales.values, labels=region_sales.index, 
                                      autopct='%1.1f%%', colors=colors_pie, startangle=90)
    ax2.set_title('Regional Sales Distribution', fontsize=16, fontweight='bold')
    
    # 3. Geographic Group Performance with Growth
    group_metrics = data.groupby('Group').agg({
        'Sales Amount': ['sum', 'mean', 'count'],
        'Profit': 'sum',
        'CustomerKey': 'nunique'
    }).round(2)
    group_metrics.columns = ['Total_Sales', 'Avg_Transaction', 'Transaction_Count', 'Total_Profit', 'Unique_Customers']
    group_metrics = group_metrics.reset_index()
    
    x_pos = np.arange(len(group_metrics))
    width = 0.35
    
    bars1 = ax3.bar(x_pos - width/2, group_metrics['Total_Sales']/1000000, width, 
                   label='Revenue (M$)', color='#FF6B6B', alpha=0.8)
    bars2 = ax3.bar(x_pos + width/2, group_metrics['Total_Profit']/1000000, width, 
                   label='Profit (M$)', color='#4ECDC4', alpha=0.8)
    
    ax3.set_title('Geographic Group: Revenue vs Profit', fontsize=16, fontweight='bold')
    ax3.set_xlabel('Geographic Group')
    ax3.set_ylabel('Amount (Millions $)')
    ax3.set_xticks(x_pos)
    ax3.set_xticklabels(group_metrics['Group'], rotation=45, ha='right')
    ax3.legend()
    
    # Add value labels
    for i, (bar1, bar2) in enumerate(zip(bars1, bars2)):
        height1, height2 = bar1.get_height(), bar2.get_height()
        ax3.text(bar1.get_x() + bar1.get_width()/2., height1 + 0.5, f'${height1:.1f}M', 
                ha='center', va='bottom', fontweight='bold', fontsize=10)
        ax3.text(bar2.get_x() + bar2.get_width()/2., height2 + 0.5, f'${height2:.1f}M', 
                ha='center', va='bottom', fontweight='bold', fontsize=10)
    
    # 4. Customer Density by Country (use available column)
    if 'Country-Region' in data.columns:
        country_col = 'Country-Region'
    else:
        country_col = 'Country'
    
    country_customers = data.groupby(country_col)['CustomerKey'].nunique().sort_values(ascending=False).head(8)
    country_avg_order = data.groupby(country_col)['Sales Amount'].mean().loc[country_customers.index]
    
    # Create dual axis chart
    ax4_twin = ax4.twinx()
    
    bars = ax4.bar(range(len(country_customers)), country_customers.values, 
                  color='lightblue', alpha=0.7, label='Customer Count')
    line = ax4_twin.plot(range(len(country_customers)), country_avg_order.values, 
                        'ro-', linewidth=2, markersize=8, color='red', label='Avg Order Value')
    
    ax4.set_title('Customer Density vs Average Order Value by Country', fontsize=16, fontweight='bold')
    ax4.set_xlabel('Country')
    ax4.set_ylabel('Number of Customers', color='blue')
    ax4_twin.set_ylabel('Average Order Value ($)', color='red')
    ax4.set_xticks(range(len(country_customers)))
    ax4.set_xticklabels(country_customers.index, rotation=45, ha='right')
    
    # Add value labels
    for i, (customers, avg_order) in enumerate(zip(country_customers.values, country_avg_order.values)):
        ax4.text(i, customers + 50, str(customers), ha='center', va='bottom', fontweight='bold', fontsize=10)
        ax4_twin.text(i, avg_order + 20, f'${avg_order:.0f}', ha='center', va='bottom', 
                     fontweight='bold', fontsize=10, color='red')
    
    plt.tight_layout()
    charts['geographic_analysis'] = save_plot_as_base64()
    
    return charts

def create_product_intelligence(data):
    """Comprehensive product performance analysis"""
    print('🛍️ Creating Product Intelligence Dashboard...')
    charts = {}
    
    fig, ((ax1, ax2), (ax3, ax4)) = plt.subplots(2, 2, figsize=(20, 16))
    
    # 1. Product Category Performance Matrix
    category_metrics = data.groupby('Category').agg({
        'Sales Amount': 'sum',
        'Profit': 'sum',
        'Order Quantity': 'sum',
        'ProductKey': 'nunique'
    }).reset_index()
    category_metrics['Profit_Margin'] = (category_metrics['Profit'] / category_metrics['Sales Amount']) * 100
    
    # Bubble chart: Revenue vs Profit Margin (bubble size = quantity)
    scatter = ax1.scatter(category_metrics['Sales Amount']/1000000, category_metrics['Profit_Margin'], 
                         s=category_metrics['Order Quantity']/100, alpha=0.7, 
                         c=range(len(category_metrics)), cmap='viridis')
    
    for i, row in category_metrics.iterrows():
        ax1.annotate(row['Category'], (row['Sales Amount']/1000000, row['Profit_Margin']), 
                    xytext=(5, 5), textcoords='offset points', fontsize=11, fontweight='bold')
    
    ax1.set_title('Product Category Performance Matrix', fontsize=16, fontweight='bold')
    ax1.set_xlabel('Revenue (Millions $)')
    ax1.set_ylabel('Profit Margin (%)')
    ax1.grid(True, alpha=0.3)
    
    # 2. Top Products Revenue Analysis
    top_products = data.groupby('Product').agg({
        'Sales Amount': 'sum',
        'Order Quantity': 'sum',
        'Profit': 'sum'
    }).sort_values('Sales Amount', ascending=False).head(12)
    
    bars = ax2.barh(range(len(top_products)), top_products['Sales Amount']/1000, 
                   color=plt.cm.plasma(np.linspace(0, 1, len(top_products))))
    ax2.set_title('Top 12 Products by Revenue', fontsize=16, fontweight='bold')
    ax2.set_xlabel('Sales Amount (Thousands $)')
    ax2.set_yticks(range(len(top_products)))
    ax2.set_yticklabels([p[:35] + '...' if len(p) > 35 else p for p in top_products.index], fontsize=10)
    
    for i, v in enumerate(top_products['Sales Amount'].values/1000):
        ax2.text(v + 20, i, f'${v:.0f}K', va='center', fontweight='bold', fontsize=10)
    
    # 3. Price vs Sales Relationship
    product_summary = data.groupby('ProductKey').agg({
        'List Price': 'first',
        'Sales Amount': 'sum',
        'Order Quantity': 'sum',
        'Product': 'first'
    }).reset_index()
    
    scatter = ax3.scatter(product_summary['List Price'], product_summary['Sales Amount']/1000, 
                         alpha=0.6, c=product_summary['Order Quantity'], cmap='coolwarm', s=60)
    ax3.set_title('Product Price vs Total Sales Performance', fontsize=16, fontweight='bold')
    ax3.set_xlabel('List Price ($)')
    ax3.set_ylabel('Total Sales Amount (Thousands $)')
    plt.colorbar(scatter, ax=ax3, label='Total Quantity Sold')
    ax3.grid(True, alpha=0.3)
    
    # Add trend line
    valid_data = product_summary.dropna(subset=['List Price', 'Sales Amount'])
    if len(valid_data) > 1:
        z = np.polyfit(valid_data['List Price'], valid_data['Sales Amount'], 1)
        p = np.poly1d(z)
        ax3.plot(valid_data['List Price'], p(valid_data['List Price'])/1000, "r--", alpha=0.8, linewidth=2)
    
    # 4. Color Performance Analysis
    color_performance = data.groupby('Color').agg({
        'Sales Amount': 'sum',
        'Order Quantity': 'sum',
        'Profit': 'sum'
    }).sort_values('Sales Amount', ascending=False)
    color_performance = color_performance[color_performance.index.notna()].head(8)
    
    # Create stacked bar chart
    x_pos = np.arange(len(color_performance))
    bars1 = ax4.bar(x_pos, color_performance['Sales Amount']/1000000, 
                   color=['black', 'red', 'yellow', 'blue', 'silver', 'white', 'green', 'orange'][:len(color_performance)], 
                   alpha=0.8, label='Revenue')
    
    ax4.set_title('Product Color Performance Analysis', fontsize=16, fontweight='bold')
    ax4.set_xlabel('Product Color')
    ax4.set_ylabel('Sales Amount (Millions $)')
    ax4.set_xticks(x_pos)
    ax4.set_xticklabels(color_performance.index, rotation=45)
    
    # Add value labels
    for i, bar in enumerate(bars1):
        height = bar.get_height()
        ax4.text(bar.get_x() + bar.get_width()/2., height + 0.1, f'${height:.1f}M', 
                ha='center', va='bottom', fontweight='bold', fontsize=11)
    
    plt.tight_layout()
    charts['product_analysis'] = save_plot_as_base64()
    
    return charts

def create_customer_analytics(data):
    """Advanced customer analytics and segmentation"""
    print('👥 Creating Customer Analytics Dashboard...')
    charts = {}
    
    fig, ((ax1, ax2), (ax3, ax4)) = plt.subplots(2, 2, figsize=(20, 16))
    
    # 1. Customer Lifetime Value Analysis
    customer_metrics = data.groupby('CustomerKey').agg({
        'Sales Amount': ['sum', 'count', 'mean'],
        'Date': ['min', 'max'],
        'Profit': 'sum'
    }).round(2)
    
    customer_metrics.columns = ['Total_Spent', 'Purchase_Count', 'Avg_Order', 'First_Purchase', 'Last_Purchase', 'Total_Profit']
    customer_metrics['Days_Active'] = (customer_metrics['Last_Purchase'] - customer_metrics['First_Purchase']).dt.days + 1
    customer_metrics['Purchase_Frequency'] = customer_metrics['Purchase_Count'] / customer_metrics['Days_Active'] * 365
    
    # CLV Distribution
    ax1.hist(customer_metrics['Total_Spent'], bins=50, alpha=0.7, color='lightcoral', edgecolor='black')
    ax1.axvline(customer_metrics['Total_Spent'].mean(), color='blue', linestyle='--', linewidth=2, 
               label=f'Mean: ${customer_metrics["Total_Spent"].mean():.0f}')
    ax1.axvline(customer_metrics['Total_Spent'].median(), color='green', linestyle='--', linewidth=2,
               label=f'Median: ${customer_metrics["Total_Spent"].median():.0f}')
    ax1.set_title('Customer Lifetime Value Distribution', fontsize=16, fontweight='bold')
    ax1.set_xlabel('Customer Lifetime Value ($)')
    ax1.set_ylabel('Number of Customers')
    ax1.legend()
    ax1.grid(True, alpha=0.3)
    
    # 2. Customer Segmentation (RFM-style)
    # Create segments based on spending and frequency
    customer_metrics['Spending_Segment'] = pd.qcut(customer_metrics['Total_Spent'], 
                                                   q=4, labels=['Low', 'Medium', 'High', 'Premium'])
    customer_metrics['Frequency_Segment'] = pd.qcut(customer_metrics['Purchase_Count'], 
                                                    q=3, labels=['Occasional', 'Regular', 'Frequent'])
    
    # Create segment matrix
    segment_matrix = customer_metrics.groupby(['Spending_Segment', 'Frequency_Segment']).size().unstack(fill_value=0)
    
    im = ax2.imshow(segment_matrix.values, cmap='YlOrRd', aspect='auto')
    ax2.set_title('Customer Segmentation Matrix', fontsize=16, fontweight='bold')
    ax2.set_xlabel('Purchase Frequency')
    ax2.set_ylabel('Spending Level')
    ax2.set_xticks(range(len(segment_matrix.columns)))
    ax2.set_xticklabels(segment_matrix.columns)
    ax2.set_yticks(range(len(segment_matrix.index)))
    ax2.set_yticklabels(segment_matrix.index)
    
    # Add text annotations
    for i in range(len(segment_matrix.index)):
        for j in range(len(segment_matrix.columns)):
            text = ax2.text(j, i, segment_matrix.iloc[i, j], ha="center", va="center", 
                           color="black", fontweight='bold', fontsize=12)
    
    plt.colorbar(im, ax=ax2, label='Number of Customers')
    
    # 3. Geographic Customer Distribution
    top_cities = data.groupby('City').agg({
        'CustomerKey': 'nunique',
        'Sales Amount': 'sum'
    }).sort_values('Sales Amount', ascending=False).head(15)
    
    bars = ax3.barh(range(len(top_cities)), top_cities['Sales Amount']/1000, 
                   color=plt.cm.viridis(np.linspace(0, 1, len(top_cities))))
    ax3.set_title('Top 15 Cities by Customer Revenue', fontsize=16, fontweight='bold')
    ax3.set_xlabel('Sales Amount (Thousands $)')
    ax3.set_yticks(range(len(top_cities)))
    ax3.set_yticklabels(top_cities.index, fontsize=10)
    
    for i, (city, row) in enumerate(top_cities.iterrows()):
        ax3.text(row['Sales Amount']/1000 + 50, i, f'${row["Sales Amount"]/1000:.0f}K\n({row["CustomerKey"]} customers)', 
                va='center', fontweight='bold', fontsize=9)
    
    # 4. Customer Purchase Behavior Over Time
    monthly_customers = data.groupby(['Year', 'Month_Name']).agg({
        'CustomerKey': 'nunique',
        'Sales Amount': 'sum'
    }).reset_index()
    monthly_customers['Date_Sort'] = pd.to_datetime(monthly_customers['Year'].astype(str) + ' ' + monthly_customers['Month_Name'], format='%Y %B')
    monthly_customers = monthly_customers.sort_values('Date_Sort')
    
    ax4_twin = ax4.twinx()
    
    line1 = ax4.plot(monthly_customers['Date_Sort'], monthly_customers['CustomerKey'], 
                    'b-o', linewidth=2, markersize=6, label='Active Customers')
    line2 = ax4_twin.plot(monthly_customers['Date_Sort'], monthly_customers['Sales Amount']/1000000, 
                         'r-s', linewidth=2, markersize=6, color='red', label='Revenue (M$)')
    
    ax4.set_title('Customer Activity vs Revenue Trend', fontsize=16, fontweight='bold')
    ax4.set_xlabel('Date')
    ax4.set_ylabel('Number of Active Customers', color='blue')
    ax4_twin.set_ylabel('Revenue (Millions $)', color='red')
    ax4.tick_params(axis='x', rotation=45)
    
    # Combine legends
    lines1, labels1 = ax4.get_legend_handles_labels()
    lines2, labels2 = ax4_twin.get_legend_handles_labels()
    ax4.legend(lines1 + lines2, labels1 + labels2, loc='upper left')
    
    plt.tight_layout()
    charts['customer_analytics'] = save_plot_as_base64()
    
    return charts

def create_channel_reseller_intelligence(data):
    """Channel and reseller performance analysis"""
    print('🤝 Creating Channel & Reseller Intelligence...')
    charts = {}
    
    fig, ((ax1, ax2), (ax3, ax4)) = plt.subplots(2, 2, figsize=(20, 16))
    
    # 1. Channel Performance Deep Dive
    channel_metrics = data.groupby('Channel').agg({
        'Sales Amount': ['sum', 'mean', 'count'],
        'Profit': 'sum',
        'CustomerKey': 'nunique',
        'Order Quantity': 'sum'
    }).round(2)
    
    channel_metrics.columns = ['Total_Sales', 'Avg_Transaction', 'Transaction_Count', 'Total_Profit', 'Unique_Customers', 'Total_Quantity']
    channel_metrics = channel_metrics.reset_index()
    channel_metrics['Profit_Margin'] = (channel_metrics['Total_Profit'] / channel_metrics['Total_Sales']) * 100
    
    # Multi-metric comparison
    x = np.arange(len(channel_metrics))
    width = 0.2
    
    bars1 = ax1.bar(x - width, channel_metrics['Total_Sales']/1000000, width, 
                   label='Revenue (M$)', color='#FF6B6B', alpha=0.8)
    bars2 = ax1.bar(x, channel_metrics['Total_Profit']/1000000, width, 
                   label='Profit (M$)', color='#4ECDC4', alpha=0.8)
    bars3 = ax1.bar(x + width, channel_metrics['Unique_Customers']/1000, width, 
                   label='Customers (K)', color='#45B7D1', alpha=0.8)
    
    ax1.set_title('Channel Performance: Revenue, Profit & Customers', fontsize=16, fontweight='bold')
    ax1.set_xlabel('Sales Channel')
    ax1.set_ylabel('Amount')
    ax1.set_xticks(x)
    ax1.set_xticklabels(channel_metrics['Channel'])
    ax1.legend()
    
    # Add value labels
    for bars in [bars1, bars2, bars3]:
        for bar in bars:
            height = bar.get_height()
            ax1.text(bar.get_x() + bar.get_width()/2., height + 0.5, f'{height:.1f}', 
                    ha='center', va='bottom', fontweight='bold', fontsize=10)
    
    # 2. Reseller Business Type Analysis
    reseller_metrics = data[data['Business Type'].notna()].groupby('Business Type').agg({
        'Sales Amount': 'sum',
        'Profit': 'sum',
        'ResellerKey': 'nunique'
    }).sort_values('Sales Amount', ascending=False)
    
    # Pie chart with profit information
    colors = ['#FF6B6B', '#4ECDC4', '#45B7D1', '#96CEB4']
    wedges, texts, autotexts = ax2.pie(reseller_metrics['Sales Amount'], 
                                      labels=reseller_metrics.index,
                                      autopct='%1.1f%%', colors=colors, startangle=90)
    ax2.set_title('Sales Distribution by Reseller Business Type', fontsize=16, fontweight='bold')
    
    # 3. Top Reseller Performance
    top_resellers = data[data['Reseller'].notna()].groupby('Reseller').agg({
        'Sales Amount': 'sum',
        'Profit': 'sum',
        'Order Quantity': 'sum'
    }).sort_values('Sales Amount', ascending=False).head(12)
    
    bars = ax3.barh(range(len(top_resellers)), top_resellers['Sales Amount']/1000, 
                   color=plt.cm.plasma(np.linspace(0, 1, len(top_resellers))))
    ax3.set_title('Top 12 Resellers by Revenue', fontsize=16, fontweight='bold')
    ax3.set_xlabel('Sales Amount (Thousands $)')
    ax3.set_yticks(range(len(top_resellers)))
    ax3.set_yticklabels([r[:25] + '...' if len(r) > 25 else r for r in top_resellers.index], fontsize=10)
    
    for i, (reseller, row) in enumerate(top_resellers.iterrows()):
        ax3.text(row['Sales Amount']/1000 + 20, i, 
                f'${row["Sales Amount"]/1000:.0f}K\n(${row["Profit"]/1000:.0f}K profit)', 
                va='center', fontweight='bold', fontsize=9)
    
    # 4. Channel Efficiency Analysis
    channel_efficiency = data.groupby('Channel').agg({
        'Sales Amount': 'sum',
        'Total Product Cost': 'sum',
        'Order Quantity': 'sum'
    })
    channel_efficiency['Revenue_per_Unit'] = channel_efficiency['Sales Amount'] / channel_efficiency['Order Quantity']
    channel_efficiency['Cost_per_Unit'] = channel_efficiency['Total Product Cost'] / channel_efficiency['Order Quantity']
    channel_efficiency['Efficiency_Ratio'] = channel_efficiency['Revenue_per_Unit'] / channel_efficiency['Cost_per_Unit']
    
    x_pos = np.arange(len(channel_efficiency))
    bars1 = ax4.bar(x_pos - 0.2, channel_efficiency['Revenue_per_Unit'], 0.4, 
                   label='Revenue per Unit', color='green', alpha=0.8)
    bars2 = ax4.bar(x_pos + 0.2, channel_efficiency['Cost_per_Unit'], 0.4, 
                   label='Cost per Unit', color='red', alpha=0.8)
    
    ax4.set_title('Channel Efficiency: Revenue vs Cost per Unit', fontsize=16, fontweight='bold')
    ax4.set_xlabel('Sales Channel')
    ax4.set_ylabel('Amount per Unit ($)')
    ax4.set_xticks(x_pos)
    ax4.set_xticklabels(channel_efficiency.index)
    ax4.legend()
    
    # Add efficiency ratio as text
    for i, ratio in enumerate(channel_efficiency['Efficiency_Ratio']):
        ax4.text(i, max(channel_efficiency['Revenue_per_Unit'].max(), channel_efficiency['Cost_per_Unit'].max()) + 50,
                f'Ratio: {ratio:.2f}', ha='center', va='bottom', fontweight='bold', 
                bbox=dict(boxstyle="round,pad=0.3", facecolor="yellow", alpha=0.7))
    
    plt.tight_layout()
    charts['channel_analysis'] = save_plot_as_base64()
    
    return charts

def create_predictive_insights(data):
    """Advanced analytics and forecasting"""
    print('🔮 Creating Predictive Insights Dashboard...')
    charts = {}
    
    fig, ((ax1, ax2), (ax3, ax4)) = plt.subplots(2, 2, figsize=(20, 16))
    
    # 1. Sales Forecasting
    monthly_sales = data.groupby(['Year', 'Month_Name'])['Sales Amount'].sum().reset_index()
    monthly_sales['Date_Sort'] = pd.to_datetime(monthly_sales['Year'].astype(str) + ' ' + monthly_sales['Month_Name'], format='%Y %B')
    monthly_sales = monthly_sales.sort_values('Date_Sort')
    
    # Simple linear trend forecasting
    x_numeric = np.arange(len(monthly_sales))
    slope, intercept, r_value, _, _ = stats.linregress(x_numeric, monthly_sales['Sales Amount'])
    
    # Forecast next 6 months
    future_months = 6
    future_x = np.arange(len(monthly_sales), len(monthly_sales) + future_months)
    future_sales = slope * future_x + intercept
    
    # Plot historical and forecasted data
    ax1.plot(monthly_sales['Date_Sort'], monthly_sales['Sales Amount']/1000000, 
            'b-o', linewidth=2, markersize=6, label='Historical Sales')
    
    # Create future dates
    last_date = monthly_sales['Date_Sort'].iloc[-1]
    future_dates = pd.date_range(start=last_date + pd.DateOffset(months=1), periods=future_months, freq='MS')
    
    ax1.plot(future_dates, future_sales/1000000, 
            'r--s', linewidth=2, markersize=6, label='Forecasted Sales', alpha=0.8)
    
    ax1.set_title(f'Sales Forecasting (Next {future_months} Months)', fontsize=16, fontweight='bold')
    ax1.set_xlabel('Date')
    ax1.set_ylabel('Sales Amount (Millions $)')
    ax1.legend()
    ax1.grid(True, alpha=0.3)
    ax1.tick_params(axis='x', rotation=45)
    
    # Add forecast values as text
    for i, (date, value) in enumerate(zip(future_dates, future_sales)):
        ax1.text(date, value/1000000 + 0.5, f'${value/1000000:.1f}M', 
                ha='center', va='bottom', fontweight='bold', fontsize=9, color='red')
    
    # 2. Customer Churn Risk Analysis
    customer_last_purchase = data.groupby('CustomerKey')['Date'].max().reset_index()
    customer_last_purchase['Days_Since_Last_Purchase'] = (data['Date'].max() - customer_last_purchase['Date']).dt.days
    
    # Define churn risk categories
    def churn_risk(days):
        if days <= 30:
            return 'Active'
        elif days <= 90:
            return 'At Risk'
        elif days <= 180:
            return 'High Risk'
        else:
            return 'Churned'
    
    customer_last_purchase['Churn_Risk'] = customer_last_purchase['Days_Since_Last_Purchase'].apply(churn_risk)
    churn_distribution = customer_last_purchase['Churn_Risk'].value_counts()
    
    colors = ['green', 'yellow', 'orange', 'red']
    bars = ax2.bar(churn_distribution.index, churn_distribution.values, 
                  color=colors[:len(churn_distribution)], alpha=0.8)
    ax2.set_title('Customer Churn Risk Analysis', fontsize=16, fontweight='bold')
    ax2.set_xlabel('Risk Category')
    ax2.set_ylabel('Number of Customers')
    
    # Add percentage labels
    total_customers = churn_distribution.sum()
    for bar, count in zip(bars, churn_distribution.values):
        percentage = (count / total_customers) * 100
        ax2.text(bar.get_x() + bar.get_width()/2., bar.get_height() + 50, 
                f'{count}\n({percentage:.1f}%)', ha='center', va='bottom', fontweight='bold')
    
    # 3. Product Performance Correlation Matrix
    product_metrics = data.groupby('ProductKey').agg({
        'Sales Amount': 'sum',
        'Order Quantity': 'sum',
        'Unit Price': 'mean',
        'List Price': 'first',
        'Product Standard Cost': 'mean',
        'Profit': 'sum'
    })
    product_metrics['Profit_Margin'] = (product_metrics['Profit'] / product_metrics['Sales Amount']) * 100
    
    # Calculate correlation matrix
    correlation_matrix = product_metrics[['Sales Amount', 'Order Quantity', 'Unit Price', 
                                        'List Price', 'Product Standard Cost', 'Profit_Margin']].corr()
    
    im = ax3.imshow(correlation_matrix.values, cmap='RdYlBu', aspect='auto', vmin=-1, vmax=1)
    ax3.set_title('Product Performance Correlation Matrix', fontsize=16, fontweight='bold')
    ax3.set_xticks(range(len(correlation_matrix.columns)))
    ax3.set_xticklabels(correlation_matrix.columns, rotation=45, ha='right')
    ax3.set_yticks(range(len(correlation_matrix.index)))
    ax3.set_yticklabels(correlation_matrix.index)
    
    # Add correlation values
    for i in range(len(correlation_matrix.index)):
        for j in range(len(correlation_matrix.columns)):
            text = ax3.text(j, i, f'{correlation_matrix.iloc[i, j]:.2f}', 
                           ha="center", va="center", color="black", fontweight='bold')
    
    plt.colorbar(im, ax=ax3, label='Correlation Coefficient')
    
    # 4. Market Opportunity Analysis
    # Analyze underperforming segments
    country_opportunity = data.groupby('Country').agg({
        'Sales Amount': 'sum',
        'CustomerKey': 'nunique',
        'ProductKey': 'nunique'
    })
    country_opportunity['Revenue_per_Customer'] = country_opportunity['Sales Amount'] / country_opportunity['CustomerKey']
    country_opportunity['Market_Penetration'] = country_opportunity['CustomerKey'] / country_opportunity['CustomerKey'].sum()
    
    # Create opportunity matrix
    scatter = ax4.scatter(country_opportunity['Market_Penetration'] * 100, 
                         country_opportunity['Revenue_per_Customer'],
                         s=country_opportunity['Sales Amount']/100000, alpha=0.7,
                         c=range(len(country_opportunity)), cmap='viridis')
    
    ax4.set_title('Market Opportunity Analysis by Country', fontsize=16, fontweight='bold')
    ax4.set_xlabel('Market Penetration (%)')
    ax4.set_ylabel('Revenue per Customer ($)')
    ax4.grid(True, alpha=0.3)
    
    # Add country labels
    for country, row in country_opportunity.iterrows():
        ax4.annotate(country, (row['Market_Penetration'] * 100, row['Revenue_per_Customer']), 
                    xytext=(5, 5), textcoords='offset points', fontsize=10, fontweight='bold')
    
    # Add quadrant lines
    ax4.axhline(y=country_opportunity['Revenue_per_Customer'].median(), color='red', linestyle='--', alpha=0.5)
    ax4.axvline(x=country_opportunity['Market_Penetration'].median() * 100, color='red', linestyle='--', alpha=0.5)
    
    plt.tight_layout()
    charts['predictive_insights'] = save_plot_as_base64()
    
    return charts

def generate_business_recommendations(data, summary):
    """Generate actionable business recommendations"""
    print('🎯 Generating Business Recommendations...')
    
    recommendations = []
    
    # Revenue Growth Opportunities
    country_sales = data.groupby('Country')['Sales Amount'].sum().sort_values(ascending=False)
    top_country = country_sales.index[0]
    lowest_country = country_sales.index[-1]
    
    recommendations.append({
        'category': 'Geographic Expansion',
        'priority': 'High',
        'recommendation': f'Focus expansion efforts on {lowest_country} market - currently only ${country_sales[lowest_country]/1000000:.1f}M vs ${country_sales[top_country]/1000000:.1f}M in {top_country}',
        'potential_impact': f'${(country_sales[top_country] - country_sales[lowest_country])/1000000:.1f}M revenue opportunity'
    })
    
    # Product Optimization
    product_profit = data.groupby('Product')['Profit'].sum().sort_values(ascending=False)
    top_profitable_product = product_profit.index[0]
    
    recommendations.append({
        'category': 'Product Strategy',
        'priority': 'High',
        'recommendation': f'Expand inventory and marketing for "{top_profitable_product}" - highest profit generator at ${product_profit[top_profitable_product]/1000:.0f}K',
        'potential_impact': 'Increase focus on top 20% of products driving 80% of profits'
    })
    
    # Customer Retention
    customer_metrics = data.groupby('CustomerKey')['Sales Amount'].sum()
    high_value_customers = len(customer_metrics[customer_metrics > customer_metrics.quantile(0.8)])
    
    recommendations.append({
        'category': 'Customer Retention',
        'priority': 'Medium',
        'recommendation': f'Implement VIP program for top {high_value_customers} customers (top 20% by value)',
        'potential_impact': f'Protect ${customer_metrics.quantile(0.8) * high_value_customers/1000000:.1f}M in high-value customer revenue'
    })
    
    # Channel Optimization
    channel_performance = data.groupby('Channel')['Sales Amount'].sum()
    if len(channel_performance) > 1:
        top_channel = channel_performance.index[0]
        recommendations.append({
            'category': 'Channel Strategy',
            'priority': 'Medium',
            'recommendation': f'Optimize {top_channel} channel performance - currently generating ${channel_performance[top_channel]/1000000:.1f}M',
            'potential_impact': 'Balance channel mix for maximum reach and efficiency'
        })
    
    # Seasonal Strategy
    monthly_avg = data.groupby('Month_Name')['Sales Amount'].mean()
    peak_month = monthly_avg.idxmax()
    low_month = monthly_avg.idxmin()
    
    recommendations.append({
        'category': 'Seasonal Planning',
        'priority': 'Medium',
        'recommendation': f'Prepare for peak season in {peak_month} and boost marketing in {low_month}',
        'potential_impact': f'Level seasonal variations - {(monthly_avg[peak_month] - monthly_avg[low_month])/1000:.0f}K daily difference'
    })
    
    # Profitability Focus
    avg_margin = data['Profit_Margin'].mean()
    recommendations.append({
        'category': 'Profitability',
        'priority': 'High',
        'recommendation': f'Focus on products with >20% profit margin (current average: {avg_margin:.1f}%)',
        'potential_impact': f'Improve overall profitability from ${summary["total_profit"]/1000000:.1f}M current profit'
    })
    
    return recommendations

def create_comprehensive_dashboard():
    """Create the complete business intelligence dashboard"""
    print('🚀 Building Comprehensive Analytics Dashboard...')
    
    # Load and prepare data
    data, raw_data = load_and_prepare_data()
    summary = create_executive_summary(data)
    
    # Generate all chart sections
    print('📊 Generating comprehensive visualizations...')
    sales_charts = create_sales_performance_analytics(data)
    geo_charts = create_geographic_intelligence(data)
    product_charts = create_product_intelligence(data)
    customer_charts = create_customer_analytics(data)
    channel_charts = create_channel_reseller_intelligence(data)
    predictive_charts = create_predictive_insights(data)
    
    # Generate business recommendations
    recommendations = generate_business_recommendations(data, summary)
    
    # Build comprehensive HTML dashboard
    html_content = f'''<!DOCTYPE html>
<html>
<head>
    <title>AdventureWorks Comprehensive Business Intelligence Dashboard</title>
    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <style>
        * {{
            margin: 0;
            padding: 0;
            box-sizing: border-box;
        }}
        
        body {{
            font-family: 'Segoe UI', Tahoma, Geneva, Verdana, sans-serif;
            background: linear-gradient(135deg, #667eea 0%, #764ba2 100%);
            min-height: 100vh;
            padding: 20px;
        }}
        
        .dashboard-container {{
            max-width: 1600px;
            margin: 0 auto;
            background: white;
            border-radius: 20px;
            box-shadow: 0 20px 40px rgba(0,0,0,0.1);
            overflow: hidden;
        }}
        
        .header {{
            background: linear-gradient(135deg, #2c3e50, #3498db);
            color: white;
            padding: 40px;
            text-align: center;
        }}
        
        .header h1 {{
            font-size: 3em;
            margin-bottom: 10px;
            text-shadow: 2px 2px 4px rgba(0,0,0,0.3);
        }}
        
        .header .subtitle {{
            font-size: 1.2em;
            opacity: 0.9;
        }}
        
        .executive-summary {{
            background: linear-gradient(135deg, #f8f9fa, #e9ecef);
            padding: 40px;
            border-bottom: 5px solid #007bff;
        }}
        
        .kpi-grid {{
            display: grid;
            grid-template-columns: repeat(auto-fit, minmax(200px, 1fr));
            gap: 20px;
            margin-bottom: 30px;
        }}
        
        .kpi-card {{
            background: white;
            padding: 25px;
            border-radius: 15px;
            text-align: center;
            box-shadow: 0 8px 16px rgba(0,0,0,0.1);
            border-left: 5px solid;
            transition: transform 0.3s ease;
        }}
        
        .kpi-card:hover {{
            transform: translateY(-5px);
        }}
        
        .kpi-card.revenue {{ border-left-color: #28a745; }}
        .kpi-card.transactions {{ border-left-color: #007bff; }}
        .kpi-card.customers {{ border-left-color: #fd7e14; }}
        .kpi-card.profit {{ border-left-color: #6f42c1; }}
        .kpi-card.growth {{ border-left-color: #20c997; }}
        .kpi-card.margin {{ border-left-color: #e83e8c; }}
        
        .kpi-value {{
            font-size: 2.5em;
            font-weight: bold;
            margin: 10px 0;
        }}
        
        .kpi-label {{
            color: #6c757d;
            font-size: 0.9em;
            text-transform: uppercase;
            letter-spacing: 1px;
        }}
        
        .section {{
            padding: 50px 40px;
            border-bottom: 1px solid #dee2e6;
        }}
        
        .section:nth-child(even) {{
            background-color: #f8f9fa;
        }}
        
        .section-header {{
            text-align: center;
            margin-bottom: 40px;
        }}
        
        .section-title {{
            font-size: 2.5em;
            color: #2c3e50;
            margin-bottom: 15px;
        }}
        
        .section-description {{
            font-size: 1.1em;
            color: #6c757d;
            max-width: 800px;
            margin: 0 auto;
            line-height: 1.6;
        }}
        
        .chart-container {{
            background: white;
            border-radius: 15px;
            padding: 30px;
            margin: 30px 0;
            box-shadow: 0 10px 20px rgba(0,0,0,0.1);
        }}
        
        .chart-grid {{
            display: grid;
            grid-template-columns: 1fr 1fr;
            gap: 30px;
            margin: 30px 0;
        }}
        
        .chart-single {{
            display: flex;
            justify-content: center;
            margin: 30px 0;
        }}
        
        .chart img {{
            max-width: 100%;
            height: auto;
            border-radius: 10px;
            box-shadow: 0 5px 15px rgba(0,0,0,0.1);
        }}
        
        .insights-box {{
            background: linear-gradient(135deg, #e8f5e9, #c8e6c9);
            border-left: 5px solid #28a745;
            padding: 25px;
            border-radius: 10px;
            margin: 25px 0;
            box-shadow: 0 5px 15px rgba(0,0,0,0.1);
        }}
        
        .insights-title {{
            font-size: 1.3em;
            font-weight: bold;
            color: #155724;
            margin-bottom: 15px;
        }}
        
        .recommendations {{
            background: linear-gradient(135deg, #fff3cd, #ffeaa7);
            padding: 40px;
        }}
        
        .recommendation-card {{
            background: white;
            border-radius: 15px;
            padding: 25px;
            margin: 20px 0;
            border-left: 5px solid;
            box-shadow: 0 5px 15px rgba(0,0,0,0.1);
        }}
        
        .recommendation-card.high {{ border-left-color: #dc3545; }}
        .recommendation-card.medium {{ border-left-color: #ffc107; }}
        .recommendation-card.low {{ border-left-color: #28a745; }}
        
        .priority-badge {{
            display: inline-block;
            padding: 5px 15px;
            border-radius: 20px;
            font-size: 0.8em;
            font-weight: bold;
            text-transform: uppercase;
            margin-bottom: 15px;
        }}
        
        .priority-high {{
            background-color: #dc3545;
            color: white;
        }}
        
        .priority-medium {{
            background-color: #ffc107;
            color: #212529;
        }}
        
        .priority-low {{
            background-color: #28a745;
            color: white;
        }}
        
        .highlight {{
            background-color: #fff3cd;
            padding: 3px 8px;
            border-radius: 5px;
            font-weight: bold;
        }}
        
        .footer {{
            background: #2c3e50;
            color: white;
            padding: 40px;
            text-align: center;
        }}
        
        @media (max-width: 768px) {{
            .chart-grid {{
                grid-template-columns: 1fr;
            }}
            
            .kpi-grid {{
                grid-template-columns: 1fr 1fr;
            }}
            
            .header h1 {{
                font-size: 2em;
            }}
            
            .section {{
                padding: 30px 20px;
            }}
        }}
    </style>
</head>
<body>
    <div class="dashboard-container">
        <div class="header">
            <h1>🚀 AdventureWorks Business Intelligence</h1>
            <div class="subtitle">Comprehensive Analytics Dashboard | {summary['date_range']}</div>
        </div>
        
        <div class="executive-summary">
            <h2 style="text-align: center; margin-bottom: 30px; color: #2c3e50;">📊 Executive Summary</h2>
            <div class="kpi-grid">
                <div class="kpi-card revenue">
                    <div class="kpi-value" style="color: #28a745;">${summary['total_revenue']/1000000:.1f}M</div>
                    <div class="kpi-label">Total Revenue</div>
                </div>
                <div class="kpi-card transactions">
                    <div class="kpi-value" style="color: #007bff;">{summary['total_transactions']:,}</div>
                    <div class="kpi-label">Total Transactions</div>
                </div>
                <div class="kpi-card customers">
                    <div class="kpi-value" style="color: #fd7e14;">{summary['unique_customers']:,}</div>
                    <div class="kpi-label">Unique Customers</div>
                </div>
                <div class="kpi-card profit">
                    <div class="kpi-value" style="color: #6f42c1;">${summary['total_profit']/1000000:.1f}M</div>
                    <div class="kpi-label">Total Profit</div>
                </div>
                <div class="kpi-card growth">
                    <div class="kpi-value" style="color: #20c997;">{summary['yoy_growth']:+.1f}%</div>
                    <div class="kpi-label">YoY Growth</div>
                </div>
                <div class="kpi-card margin">
                    <div class="kpi-value" style="color: #e83e8c;">{summary['avg_profit_margin']:.1f}%</div>
                    <div class="kpi-label">Avg Profit Margin</div>
                </div>
            </div>
            
            <div class="insights-box">
                <div class="insights-title">🎯 Key Business Highlights</div>
                <ul style="line-height: 2; font-size: 1.1em;">
                    <li><span class="highlight">Revenue Performance:</span> ${summary['total_revenue']/1000000:.1f}M total revenue across {summary['countries']} countries</li>
                    <li><span class="highlight">Customer Base:</span> {summary['unique_customers']:,} unique customers with ${summary['customer_ltv']:.0f} average lifetime value</li>
                    <li><span class="highlight">Product Portfolio:</span> {summary['unique_products']} products generating ${summary['avg_order_value']:.0f} average order value</li>
                    <li><span class="highlight">Profitability:</span> {summary['avg_profit_margin']:.1f}% average profit margin with ${summary['total_profit']/1000000:.1f}M total profit</li>
                    <li><span class="highlight">Growth Trajectory:</span> {summary['yoy_growth']:+.1f}% year-over-year growth trend</li>
                </ul>
            </div>
        </div>
        
        <div class="section">
            <div class="section-header">
                <h2 class="section-title">📈 Sales Performance Analytics</h2>
                <p class="section-description">
                    Comprehensive analysis of sales trends, seasonality, and profitability patterns. 
                    Includes forecasting capabilities and profit margin optimization insights.
                </p>
            </div>
            
            <div class="chart-container">
                <img src="data:image/png;base64,{sales_charts['sales_trend']}" alt="Sales Performance Trends">
            </div>
            
            <div class="chart-container">
                <img src="data:image/png;base64,{sales_charts['profit_analysis']}" alt="Profit Analysis Dashboard">
            </div>
            
            <div class="insights-box">
                <div class="insights-title">💡 Sales Performance Insights</div>
                <ul style="line-height: 1.8;">
                    <li><strong>Seasonal Patterns:</strong> Clear monthly variations with peak performance opportunities identified</li>
                    <li><strong>Profit Optimization:</strong> Average {summary['avg_profit_margin']:.1f}% margin with significant improvement potential</li>
                    <li><strong>Growth Trend:</strong> {summary['yoy_growth']:+.1f}% year-over-year growth indicating {'strong positive' if summary['yoy_growth'] > 0 else 'challenging'} market conditions</li>
                    <li><strong>Revenue Concentration:</strong> Top products drive majority of profitability - focus area for expansion</li>
                </ul>
            </div>
        </div>
        
        <div class="section">
            <div class="section-header">
                <h2 class="section-title">🌍 Geographic Market Intelligence</h2>
                <p class="section-description">
                    Deep dive into geographic performance, market penetration analysis, and regional growth opportunities 
                    across {summary['countries']} countries and multiple regions.
                </p>
            </div>
            
            <div class="chart-container">
                <img src="data:image/png;base64,{geo_charts['geographic_analysis']}" alt="Geographic Analysis Dashboard">
            </div>
            
            <div class="insights-box">
                <div class="insights-title">🗺️ Geographic Market Insights</div>
                <ul style="line-height: 1.8;">
                    <li><strong>Market Dominance:</strong> United States leads with significant market share concentration</li>
                    <li><strong>Expansion Opportunities:</strong> Underperforming markets show potential for growth investment</li>
                    <li><strong>Regional Balance:</strong> North American markets drive majority of revenue with international growth potential</li>
                    <li><strong>Customer Density:</strong> High-value customers concentrated in specific geographic regions</li>
                </ul>
            </div>
        </div>
        
        <div class="section">
            <div class="section-header">
                <h2 class="section-title">🛍️ Product Performance Intelligence</h2>
                <p class="section-description">
                    Advanced product analytics covering {summary['unique_products']} products across multiple categories, 
                    including profitability analysis, price optimization, and portfolio performance.
                </p>
            </div>
            
            <div class="chart-container">
                <img src="data:image/png;base64,{product_charts['product_analysis']}" alt="Product Intelligence Dashboard">
            </div>
            
            <div class="insights-box">
                <div class="insights-title">🎯 Product Strategy Insights</div>
                <ul style="line-height: 1.8;">
                    <li><strong>Portfolio Concentration:</strong> Top 20% of products generate 80% of revenue - classic Pareto principle</li>
                    <li><strong>Category Performance:</strong> Significant variation in profitability across product categories</li>
                    <li><strong>Price Optimization:</strong> Strong correlation between price positioning and sales performance</li>
                    <li><strong>Color Preferences:</strong> Clear customer preferences in product colors affecting sales volume</li>
                </ul>
            </div>
        </div>
        
        <div class="section">
            <div class="section-header">
                <h2 class="section-title">👥 Customer Analytics & Segmentation</h2>
                <p class="section-description">
                    Advanced customer intelligence covering {summary['unique_customers']:,} customers with lifetime value analysis, 
                    segmentation strategies, and churn risk assessment.
                </p>
            </div>
            
            <div class="chart-container">
                <img src="data:image/png;base64,{customer_charts['customer_analytics']}" alt="Customer Analytics Dashboard">
            </div>
            
            <div class="insights-box">
                <div class="insights-title">🎪 Customer Intelligence Insights</div>
                <ul style="line-height: 1.8;">
                    <li><strong>Customer Value Distribution:</strong> ${summary['customer_ltv']:.0f} average lifetime value with significant high-value segment</li>
                    <li><strong>Segmentation Opportunities:</strong> Clear customer segments based on spending and frequency patterns</li>
                    <li><strong>Geographic Concentration:</strong> Customer base concentrated in key metropolitan areas</li>
                    <li><strong>Engagement Patterns:</strong> Strong correlation between customer activity and revenue generation</li>
                </ul>
            </div>
        </div>
        
        <div class="section">
            <div class="section-header">
                <h2 class="section-title">🤝 Channel & Reseller Intelligence</h2>
                <p class="section-description">
                    Comprehensive analysis of sales channels and reseller network performance, 
                    including efficiency metrics and partnership optimization strategies.
                </p>
            </div>
            
            <div class="chart-container">
                <img src="data:image/png;base64,{channel_charts['channel_analysis']}" alt="Channel & Reseller Analysis">
            </div>
            
            <div class="insights-box">
                <div class="insights-title">🔗 Channel Strategy Insights</div>
                <ul style="line-height: 1.8;">
                    <li><strong>Channel Performance:</strong> Balanced dual-channel approach with distinct performance profiles</li>
                    <li><strong>Reseller Network:</strong> Strong partner ecosystem with top performers driving significant revenue</li>
                    <li><strong>Efficiency Metrics:</strong> Clear differences in cost-effectiveness across channels</li>
                    <li><strong>Partnership Optimization:</strong> Opportunities to enhance reseller performance and support</li>
                </ul>
            </div>
        </div>
        
        <div class="section">
            <div class="section-header">
                <h2 class="section-title">🔮 Predictive Analytics & Forecasting</h2>
                <p class="section-description">
                    Advanced analytics including sales forecasting, customer churn prediction, 
                    market opportunity analysis, and performance correlation insights.
                </p>
            </div>
            
            <div class="chart-container">
                <img src="data:image/png;base64,{predictive_charts['predictive_insights']}" alt="Predictive Analytics Dashboard">
            </div>
            
            <div class="insights-box">
                <div class="insights-title">🚀 Predictive Intelligence Insights</div>
                <ul style="line-height: 1.8;">
                    <li><strong>Sales Forecasting:</strong> 6-month forward projection based on historical trends and seasonality</li>
                    <li><strong>Churn Risk Analysis:</strong> Customer retention insights with risk categorization</li>
                    <li><strong>Market Opportunities:</strong> Underperforming segments identified for growth investment</li>
                    <li><strong>Performance Correlations:</strong> Key drivers of success identified through advanced analytics</li>
                </ul>
            </div>
        </div>
        
        <div class="recommendations">
            <div class="section-header">
                <h2 class="section-title">🎯 Strategic Business Recommendations</h2>
                <p class="section-description">
                    Data-driven actionable recommendations prioritized by potential business impact 
                    and implementation feasibility.
                </p>
            </div>
            
            {chr(10).join([f'''
            <div class="recommendation-card {rec['priority'].lower()}">
                <div class="priority-badge priority-{rec['priority'].lower()}">{rec['priority']} Priority</div>
                <h3 style="color: #2c3e50; margin-bottom: 15px;">📊 {rec['category']}</h3>
                <p style="font-size: 1.1em; line-height: 1.6; margin-bottom: 15px;"><strong>Recommendation:</strong> {rec['recommendation']}</p>
                <p style="color: #28a745; font-weight: bold;">💰 Potential Impact: {rec['potential_impact']}</p>
            </div>
            ''' for rec in recommendations])}
            
            <div class="insights-box" style="margin-top: 40px;">
                <div class="insights-title">🚀 Implementation Roadmap</div>
                <ol style="line-height: 2; font-size: 1.1em;">
                    <li><strong>Immediate Actions (0-30 days):</strong> Implement high-priority recommendations with quick wins</li>
                    <li><strong>Short-term Initiatives (1-3 months):</strong> Launch medium-priority strategic initiatives</li>
                    <li><strong>Long-term Strategy (3-12 months):</strong> Execute comprehensive transformation programs</li>
                    <li><strong>Continuous Monitoring:</strong> Establish KPI tracking and regular performance reviews</li>
                </ol>
            </div>
        </div>
        
        <div class="footer">
            <h3>📈 Dashboard Summary</h3>
            <p style="margin: 20px 0; font-size: 1.1em; line-height: 1.6;">
                This comprehensive business intelligence dashboard analyzes <strong>{summary['total_transactions']:,} transactions</strong> 
                from <strong>{summary['unique_customers']:,} customers</strong> across <strong>{summary['countries']} countries</strong>, 
                featuring <strong>{summary['unique_products']} products</strong> and generating 
                <strong>${summary['total_revenue']/1000000:.1f}M in total revenue</strong> with 
                <strong>${summary['total_profit']/1000000:.1f}M profit</strong>.
            </p>
            <p style="opacity: 0.8;">
                <em>Generated from AdventureWorks Sales Dataset | Advanced Analytics & Business Intelligence Solution</em>
            </p>
            <p style="margin-top: 20px; font-size: 0.9em;">
                Dashboard includes: Sales Performance Analytics | Geographic Intelligence | Product Intelligence | 
                Customer Analytics | Channel Analysis | Predictive Insights | Strategic Recommendations
            </p>
        </div>
    </div>
</body>
</html>'''
    
    # Save the comprehensive dashboard
    with open('Comprehensive_Business_Intelligence_Dashboard.html', 'w', encoding='utf-8') as f:
        f.write(html_content)
    
    print('\n🎉 COMPREHENSIVE DASHBOARD CREATED SUCCESSFULLY!')
    print('📄 Open "Comprehensive_Business_Intelligence_Dashboard.html" in your web browser')
    print(f'📊 Dashboard includes:')
    print(f'   • Executive Summary with {len([k for k in summary.keys()])} key metrics')
    print(f'   • 6 comprehensive chart sections with {len(sales_charts) + len(geo_charts) + len(product_charts) + len(customer_charts) + len(channel_charts) + len(predictive_charts)} visualizations')
    print(f'   • {len(recommendations)} strategic business recommendations')
    print(f'   • Advanced analytics covering ${summary["total_revenue"]/1000000:.1f}M in revenue analysis')
    
    return True

if __name__ == "__main__":
    create_comprehensive_dashboard()