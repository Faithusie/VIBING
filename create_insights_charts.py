#!/usr/bin/env python3
import pandas as pd
import numpy as np
import matplotlib.pyplot as plt
import seaborn as sns
from datetime import datetime
import warnings
warnings.filterwarnings('ignore')

# Set style for better-looking plots
plt.style.use('seaborn-v0_8')
sns.set_palette("husl")

def load_data():
    """Load all data sheets from the Excel file"""
    file_path = "AdventureWorks Sales (1).xlsx"
    
    # Read all sheets
    excel_file = pd.ExcelFile(file_path)
    data = {}
    
    for sheet in excel_file.sheet_names:
        data[sheet] = pd.read_excel(file_path, sheet_name=sheet)
    
    return data

def create_sales_performance_charts(data):
    """Create sales performance visualizations"""
    sales_data = data['Sales_data']
    date_data = data['Date_data']
    
    # Merge sales with date data
    sales_with_dates = sales_data.merge(date_data, left_on='OrderDateKey', right_on='DateKey', how='left')
    
    fig, axes = plt.subplots(2, 2, figsize=(20, 16))
    fig.suptitle('📊 Sales Performance Analytics', fontsize=20, fontweight='bold', y=0.98)
    
    # 1. Monthly Sales Trend
    monthly_sales = sales_with_dates.groupby('Month')['Sales Amount'].agg(['sum', 'count']).reset_index()
    monthly_sales['Month_Sort'] = pd.to_datetime(monthly_sales['Month'], format='%Y %b').dt.strftime('%m-%Y')
    monthly_sales = monthly_sales.sort_values('Month_Sort')
    
    axes[0,0].plot(range(len(monthly_sales)), monthly_sales['sum']/1000, marker='o', linewidth=3, markersize=8)
    axes[0,0].set_title('Monthly Sales Revenue Trend', fontsize=14, fontweight='bold')
    axes[0,0].set_xlabel('Month')
    axes[0,0].set_ylabel('Sales Amount (Thousands $)')
    axes[0,0].tick_params(axis='x', rotation=45)
    axes[0,0].set_xticks(range(0, len(monthly_sales), 6))
    axes[0,0].set_xticklabels(monthly_sales['Month'].iloc[::6], rotation=45)
    axes[0,0].grid(True, alpha=0.3)
    
    # 2. Sales by Fiscal Year
    yearly_sales = sales_with_dates.groupby('Fiscal Year')['Sales Amount'].sum().reset_index()
    yearly_sales = yearly_sales.sort_values('Fiscal Year')
    
    bars = axes[0,1].bar(yearly_sales['Fiscal Year'], yearly_sales['Sales Amount']/1000000, 
                        color=['#FF6B6B', '#4ECDC4', '#45B7D1', '#96CEB4'])
    axes[0,1].set_title('Annual Sales Revenue by Fiscal Year', fontsize=14, fontweight='bold')
    axes[0,1].set_xlabel('Fiscal Year')
    axes[0,1].set_ylabel('Sales Amount (Millions $)')
    
    # Add value labels on bars
    for bar in bars:
        height = bar.get_height()
        axes[0,1].text(bar.get_x() + bar.get_width()/2., height + 0.5,
                      f'${height:.1f}M', ha='center', va='bottom', fontweight='bold')
    
    # 3. Profit Margin Analysis
    sales_data['Profit'] = sales_data['Sales Amount'] - sales_data['Total Product Cost']
    sales_data['Profit_Margin'] = (sales_data['Profit'] / sales_data['Sales Amount']) * 100
    
    axes[1,0].hist(sales_data['Profit_Margin'], bins=50, alpha=0.7, color='skyblue', edgecolor='black')
    axes[1,0].axvline(sales_data['Profit_Margin'].mean(), color='red', linestyle='--', 
                     linewidth=2, label=f'Mean: {sales_data["Profit_Margin"].mean():.1f}%')
    axes[1,0].set_title('Distribution of Profit Margins', fontsize=14, fontweight='bold')
    axes[1,0].set_xlabel('Profit Margin (%)')
    axes[1,0].set_ylabel('Frequency')
    axes[1,0].legend()
    axes[1,0].grid(True, alpha=0.3)
    
    # 4. Sales Amount vs Order Quantity
    sample_data = sales_data.sample(n=min(5000, len(sales_data)))  # Sample for better visualization
    scatter = axes[1,1].scatter(sample_data['Order Quantity'], sample_data['Sales Amount'], 
                               alpha=0.6, c=sample_data['Unit Price'], cmap='viridis')
    axes[1,1].set_title('Sales Amount vs Order Quantity', fontsize=14, fontweight='bold')
    axes[1,1].set_xlabel('Order Quantity')
    axes[1,1].set_ylabel('Sales Amount ($)')
    plt.colorbar(scatter, ax=axes[1,1], label='Unit Price ($)')
    axes[1,1].grid(True, alpha=0.3)
    
    plt.tight_layout()
    plt.savefig('sales_performance_charts.png', dpi=300, bbox_inches='tight')
    plt.show()

def create_geographic_charts(data):
    """Create geographic analysis visualizations"""
    sales_data = data['Sales_data']
    territory_data = data['Sales Territory_data']
    
    # Merge sales with territory data
    sales_territory = sales_data.merge(territory_data, on='SalesTerritoryKey', how='left')
    
    fig, axes = plt.subplots(2, 2, figsize=(20, 16))
    fig.suptitle('🌍 Geographic Performance Analysis', fontsize=20, fontweight='bold', y=0.98)
    
    # 1. Sales by Country
    country_sales = sales_territory.groupby('Country')['Sales Amount'].sum().sort_values(ascending=False)
    
    axes[0,0].bar(range(len(country_sales)), country_sales.values/1000000, 
                 color=plt.cm.Set3(np.linspace(0, 1, len(country_sales))))
    axes[0,0].set_title('Sales Revenue by Country', fontsize=14, fontweight='bold')
    axes[0,0].set_xlabel('Country')
    axes[0,0].set_ylabel('Sales Amount (Millions $)')
    axes[0,0].set_xticks(range(len(country_sales)))
    axes[0,0].set_xticklabels(country_sales.index, rotation=45, ha='right')
    
    # Add value labels
    for i, v in enumerate(country_sales.values/1000000):
        axes[0,0].text(i, v + 0.5, f'${v:.1f}M', ha='center', va='bottom', fontweight='bold')
    
    # 2. Sales by Region (US regions)
    region_sales = sales_territory.groupby('Region')['Sales Amount'].sum().sort_values(ascending=False)
    
    colors = plt.cm.Paired(np.linspace(0, 1, len(region_sales)))
    wedges, texts, autotexts = axes[0,1].pie(region_sales.values, labels=region_sales.index, 
                                            autopct='%1.1f%%', colors=colors, startangle=90)
    axes[0,1].set_title('Sales Distribution by Region', fontsize=14, fontweight='bold')
    
    # 3. Geographic Group Performance
    group_sales = sales_territory.groupby('Group')['Sales Amount'].agg(['sum', 'mean', 'count']).reset_index()
    
    x_pos = np.arange(len(group_sales))
    bars = axes[1,0].bar(x_pos, group_sales['sum']/1000000, color=['#FF6B6B', '#4ECDC4', '#45B7D1', '#96CEB4'])
    axes[1,0].set_title('Sales Revenue by Geographic Group', fontsize=14, fontweight='bold')
    axes[1,0].set_xlabel('Geographic Group')
    axes[1,0].set_ylabel('Sales Amount (Millions $)')
    axes[1,0].set_xticks(x_pos)
    axes[1,0].set_xticklabels(group_sales['Group'], rotation=45, ha='right')
    
    # Add value labels
    for i, bar in enumerate(bars):
        height = bar.get_height()
        axes[1,0].text(bar.get_x() + bar.get_width()/2., height + 0.5,
                      f'${height:.1f}M', ha='center', va='bottom', fontweight='bold')
    
    # 4. Territory Performance Heatmap
    territory_metrics = sales_territory.groupby(['Country', 'Region']).agg({
        'Sales Amount': 'sum',
        'Order Quantity': 'sum',
        'Sales Amount': 'count'
    }).reset_index()
    
    # Create a pivot table for heatmap
    pivot_data = sales_territory.groupby(['Group', 'Country'])['Sales Amount'].sum().unstack(fill_value=0)
    
    sns.heatmap(pivot_data/1000000, annot=True, fmt='.1f', cmap='YlOrRd', 
                ax=axes[1,1], cbar_kws={'label': 'Sales Amount (Millions $)'})
    axes[1,1].set_title('Sales Heatmap: Group vs Country', fontsize=14, fontweight='bold')
    axes[1,1].set_xlabel('Country')
    axes[1,1].set_ylabel('Geographic Group')
    
    plt.tight_layout()
    plt.savefig('geographic_analysis_charts.png', dpi=300, bbox_inches='tight')
    plt.show()

def create_product_charts(data):
    """Create product analysis visualizations"""
    sales_data = data['Sales_data']
    product_data = data['Product_data']
    
    # Merge sales with product data
    sales_products = sales_data.merge(product_data, on='ProductKey', how='left')
    
    fig, axes = plt.subplots(2, 2, figsize=(20, 16))
    fig.suptitle('🛍️ Product Performance Analysis', fontsize=20, fontweight='bold', y=0.98)
    
    # 1. Top 10 Products by Revenue
    top_products = sales_products.groupby('Product')['Sales Amount'].sum().sort_values(ascending=False).head(10)
    
    axes[0,0].barh(range(len(top_products)), top_products.values/1000, 
                  color=plt.cm.viridis(np.linspace(0, 1, len(top_products))))
    axes[0,0].set_title('Top 10 Products by Revenue', fontsize=14, fontweight='bold')
    axes[0,0].set_xlabel('Sales Amount (Thousands $)')
    axes[0,0].set_ylabel('Product')
    axes[0,0].set_yticks(range(len(top_products)))
    axes[0,0].set_yticklabels([p[:30] + '...' if len(p) > 30 else p for p in top_products.index])
    
    # Add value labels
    for i, v in enumerate(top_products.values/1000):
        axes[0,0].text(v + 10, i, f'${v:.0f}K', va='center', fontweight='bold')
    
    # 2. Sales by Category
    category_sales = sales_products.groupby('Category')['Sales Amount'].sum().sort_values(ascending=False)
    
    colors = ['#FF6B6B', '#4ECDC4', '#45B7D1', '#96CEB4', '#FECA57']
    wedges, texts, autotexts = axes[0,1].pie(category_sales.values, labels=category_sales.index, 
                                            autopct='%1.1f%%', colors=colors[:len(category_sales)], 
                                            startangle=90)
    axes[0,1].set_title('Sales Distribution by Product Category', fontsize=14, fontweight='bold')
    
    # 3. Price vs Sales Relationship
    product_summary = sales_products.groupby('ProductKey').agg({
        'List Price': 'first',
        'Sales Amount': 'sum',
        'Order Quantity': 'sum'
    }).reset_index()
    
    scatter = axes[1,0].scatter(product_summary['List Price'], product_summary['Sales Amount']/1000, 
                               alpha=0.6, c=product_summary['Order Quantity'], cmap='plasma', s=60)
    axes[1,0].set_title('Product Price vs Total Sales', fontsize=14, fontweight='bold')
    axes[1,0].set_xlabel('List Price ($)')
    axes[1,0].set_ylabel('Total Sales Amount (Thousands $)')
    plt.colorbar(scatter, ax=axes[1,0], label='Total Quantity Sold')
    axes[1,0].grid(True, alpha=0.3)
    
    # 4. Top Colors Performance
    color_sales = sales_products.groupby('Color')['Sales Amount'].sum().sort_values(ascending=False).head(8)
    color_sales = color_sales[color_sales.index.notna()]  # Remove NaN colors
    
    bars = axes[1,1].bar(range(len(color_sales)), color_sales.values/1000000, 
                        color=['black', 'red', 'yellow', 'blue', 'silver', 'white', 'green', 'orange'][:len(color_sales)])
    axes[1,1].set_title('Sales Performance by Product Color', fontsize=14, fontweight='bold')
    axes[1,1].set_xlabel('Color')
    axes[1,1].set_ylabel('Sales Amount (Millions $)')
    axes[1,1].set_xticks(range(len(color_sales)))
    axes[1,1].set_xticklabels(color_sales.index, rotation=45)
    
    # Add value labels
    for i, bar in enumerate(bars):
        height = bar.get_height()
        axes[1,1].text(bar.get_x() + bar.get_width()/2., height + 0.02,
                      f'${height:.1f}M', ha='center', va='bottom', fontweight='bold')
    
    plt.tight_layout()
    plt.savefig('product_analysis_charts.png', dpi=300, bbox_inches='tight')
    plt.show()

def create_customer_charts(data):
    """Create customer analysis visualizations"""
    sales_data = data['Sales_data']
    customer_data = data['Customer_data']
    
    # Merge sales with customer data
    sales_customers = sales_data.merge(customer_data, on='CustomerKey', how='left')
    
    fig, axes = plt.subplots(2, 2, figsize=(20, 16))
    fig.suptitle('👥 Customer Analytics Dashboard', fontsize=20, fontweight='bold', y=0.98)
    
    # 1. Customer Value Distribution
    customer_value = sales_customers.groupby('CustomerKey')['Sales Amount'].sum()
    
    axes[0,0].hist(customer_value, bins=50, alpha=0.7, color='lightcoral', edgecolor='black')
    axes[0,0].axvline(customer_value.mean(), color='blue', linestyle='--', 
                     linewidth=2, label=f'Mean: ${customer_value.mean():.0f}')
    axes[0,0].axvline(customer_value.median(), color='green', linestyle='--', 
                     linewidth=2, label=f'Median: ${customer_value.median():.0f}')
    axes[0,0].set_title('Customer Lifetime Value Distribution', fontsize=14, fontweight='bold')
    axes[0,0].set_xlabel('Customer Lifetime Value ($)')
    axes[0,0].set_ylabel('Number of Customers')
    axes[0,0].legend()
    axes[0,0].grid(True, alpha=0.3)
    
    # 2. Top Countries by Customer Count
    country_customers = sales_customers.groupby('Country-Region')['CustomerKey'].nunique().sort_values(ascending=False).head(10)
    
    axes[0,1].bar(range(len(country_customers)), country_customers.values, 
                 color=plt.cm.Set2(np.linspace(0, 1, len(country_customers))))
    axes[0,1].set_title('Customer Distribution by Country', fontsize=14, fontweight='bold')
    axes[0,1].set_xlabel('Country')
    axes[0,1].set_ylabel('Number of Unique Customers')
    axes[0,1].set_xticks(range(len(country_customers)))
    axes[0,1].set_xticklabels(country_customers.index, rotation=45, ha='right')
    
    # Add value labels
    for i, v in enumerate(country_customers.values):
        axes[0,1].text(i, v + 50, str(v), ha='center', va='bottom', fontweight='bold')
    
    # 3. Customer Segmentation by Purchase Frequency
    customer_frequency = sales_customers.groupby('CustomerKey').size()
    
    # Create segments
    segments = []
    segment_labels = []
    for freq in customer_frequency:
        if freq == 1:
            segments.append('One-time')
        elif freq <= 3:
            segments.append('Occasional (2-3)')
        elif freq <= 10:
            segments.append('Regular (4-10)')
        else:
            segments.append('Frequent (10+)')
    
    segment_counts = pd.Series(segments).value_counts()
    
    colors = ['#FF6B6B', '#4ECDC4', '#45B7D1', '#96CEB4']
    wedges, texts, autotexts = axes[1,0].pie(segment_counts.values, labels=segment_counts.index, 
                                            autopct='%1.1f%%', colors=colors, startangle=90)
    axes[1,0].set_title('Customer Segmentation by Purchase Frequency', fontsize=14, fontweight='bold')
    
    # 4. Top 15 Cities by Sales
    city_sales = sales_customers.groupby('City')['Sales Amount'].sum().sort_values(ascending=False).head(15)
    
    axes[1,1].barh(range(len(city_sales)), city_sales.values/1000, 
                  color=plt.cm.plasma(np.linspace(0, 1, len(city_sales))))
    axes[1,1].set_title('Top 15 Cities by Sales Revenue', fontsize=14, fontweight='bold')
    axes[1,1].set_xlabel('Sales Amount (Thousands $)')
    axes[1,1].set_ylabel('City')
    axes[1,1].set_yticks(range(len(city_sales)))
    axes[1,1].set_yticklabels(city_sales.index)
    
    # Add value labels
    for i, v in enumerate(city_sales.values/1000):
        axes[1,1].text(v + 20, i, f'${v:.0f}K', va='center', fontweight='bold')
    
    plt.tight_layout()
    plt.savefig('customer_analysis_charts.png', dpi=300, bbox_inches='tight')
    plt.show()

def create_channel_reseller_charts(data):
    """Create channel and reseller analysis visualizations"""
    sales_order_data = data['Sales Order_data']
    sales_data = data['Sales_data']
    reseller_data = data['Reseller_data']
    
    # Merge data
    sales_with_orders = sales_data.merge(sales_order_data, on='SalesOrderLineKey', how='left')
    sales_with_resellers = sales_with_orders.merge(reseller_data, on='ResellerKey', how='left')
    
    fig, axes = plt.subplots(2, 2, figsize=(20, 16))
    fig.suptitle('🤝 Sales Channel & Reseller Performance', fontsize=20, fontweight='bold', y=0.98)
    
    # 1. Channel Performance Comparison
    channel_performance = sales_with_orders.groupby('Channel').agg({
        'Sales Amount': ['sum', 'count', 'mean']
    }).round(2)
    
    channel_performance.columns = ['Total_Sales', 'Transaction_Count', 'Avg_Transaction']
    channel_performance = channel_performance.reset_index()
    
    x = np.arange(len(channel_performance))
    width = 0.35
    
    bars1 = axes[0,0].bar(x - width/2, channel_performance['Total_Sales']/1000000, width, 
                         label='Total Sales (Millions $)', color='skyblue')
    
    ax2 = axes[0,0].twinx()
    bars2 = ax2.bar(x + width/2, channel_performance['Transaction_Count']/1000, width, 
                   label='Transactions (Thousands)', color='lightcoral')
    
    axes[0,0].set_title('Sales Channel Performance Comparison', fontsize=14, fontweight='bold')
    axes[0,0].set_xlabel('Sales Channel')
    axes[0,0].set_ylabel('Total Sales (Millions $)', color='blue')
    ax2.set_ylabel('Number of Transactions (Thousands)', color='red')
    axes[0,0].set_xticks(x)
    axes[0,0].set_xticklabels(channel_performance['Channel'])
    
    # Add value labels
    for i, (bar1, bar2) in enumerate(zip(bars1, bars2)):
        height1 = bar1.get_height()
        height2 = bar2.get_height()
        axes[0,0].text(bar1.get_x() + bar1.get_width()/2., height1 + 1,
                      f'${height1:.1f}M', ha='center', va='bottom', fontweight='bold')
        ax2.text(bar2.get_x() + bar2.get_width()/2., height2 + 1,
                f'{height2:.1f}K', ha='center', va='bottom', fontweight='bold')
    
    # 2. Reseller Business Type Performance
    business_type_sales = sales_with_resellers.groupby('Business Type')['Sales Amount'].sum().sort_values(ascending=False)
    business_type_sales = business_type_sales[business_type_sales.index.notna()]
    
    colors = ['#FF6B6B', '#4ECDC4', '#45B7D1', '#96CEB4']
    wedges, texts, autotexts = axes[0,1].pie(business_type_sales.values, labels=business_type_sales.index, 
                                            autopct='%1.1f%%', colors=colors[:len(business_type_sales)], 
                                            startangle=90)
    axes[0,1].set_title('Sales Distribution by Reseller Business Type', fontsize=14, fontweight='bold')
    
    # 3. Top 10 Resellers by Sales
    top_resellers = sales_with_resellers.groupby('Reseller')['Sales Amount'].sum().sort_values(ascending=False).head(10)
    top_resellers = top_resellers[top_resellers.index.notna()]
    
    axes[1,0].barh(range(len(top_resellers)), top_resellers.values/1000, 
                  color=plt.cm.viridis(np.linspace(0, 1, len(top_resellers))))
    axes[1,0].set_title('Top 10 Resellers by Sales Revenue', fontsize=14, fontweight='bold')
    axes[1,0].set_xlabel('Sales Amount (Thousands $)')
    axes[1,0].set_ylabel('Reseller')
    axes[1,0].set_yticks(range(len(top_resellers)))
    axes[1,0].set_yticklabels([r[:30] + '...' if len(r) > 30 else r for r in top_resellers.index])
    
    # Add value labels
    for i, v in enumerate(top_resellers.values/1000):
        axes[1,0].text(v + 20, i, f'${v:.0f}K', va='center', fontweight='bold')
    
    # 4. Average Transaction Value by Channel
    avg_transaction = sales_with_orders.groupby('Channel')['Sales Amount'].mean()
    
    bars = axes[1,1].bar(range(len(avg_transaction)), avg_transaction.values, 
                        color=['#FF6B6B', '#4ECDC4'])
    axes[1,1].set_title('Average Transaction Value by Channel', fontsize=14, fontweight='bold')
    axes[1,1].set_xlabel('Sales Channel')
    axes[1,1].set_ylabel('Average Transaction Value ($)')
    axes[1,1].set_xticks(range(len(avg_transaction)))
    axes[1,1].set_xticklabels(avg_transaction.index)
    
    # Add value labels
    for i, bar in enumerate(bars):
        height = bar.get_height()
        axes[1,1].text(bar.get_x() + bar.get_width()/2., height + 10,
                      f'${height:.0f}', ha='center', va='bottom', fontweight='bold')
    
    plt.tight_layout()
    plt.savefig('channel_reseller_charts.png', dpi=300, bbox_inches='tight')
    plt.show()

def create_time_series_charts(data):
    """Create time series analysis visualizations"""
    sales_data = data['Sales_data']
    date_data = data['Date_data']
    
    # Merge sales with date data
    sales_with_dates = sales_data.merge(date_data, left_on='OrderDateKey', right_on='DateKey', how='left')
    sales_with_dates['Date'] = pd.to_datetime(sales_with_dates['Date'])
    
    fig, axes = plt.subplots(2, 2, figsize=(20, 16))
    fig.suptitle('📈 Time Series Analysis Dashboard', fontsize=20, fontweight='bold', y=0.98)
    
    # 1. Daily Sales Trend
    daily_sales = sales_with_dates.groupby('Date')['Sales Amount'].sum().reset_index()
    daily_sales = daily_sales.sort_values('Date')
    
    axes[0,0].plot(daily_sales['Date'], daily_sales['Sales Amount']/1000, linewidth=2, color='blue')
    axes[0,0].set_title('Daily Sales Revenue Trend', fontsize=14, fontweight='bold')
    axes[0,0].set_xlabel('Date')
    axes[0,0].set_ylabel('Sales Amount (Thousands $)')
    axes[0,0].tick_params(axis='x', rotation=45)
    axes[0,0].grid(True, alpha=0.3)
    
    # Add trend line
    from scipy import stats
    x_numeric = np.arange(len(daily_sales))
    slope, intercept, r_value, p_value, std_err = stats.linregress(x_numeric, daily_sales['Sales Amount'])
    trend_line = slope * x_numeric + intercept
    axes[0,0].plot(daily_sales['Date'], trend_line/1000, '--', color='red', alpha=0.8, 
                  label=f'Trend (R²={r_value**2:.3f})')
    axes[0,0].legend()
    
    # 2. Seasonal Patterns (Monthly)
    sales_with_dates['Month_Name'] = sales_with_dates['Date'].dt.strftime('%B')
    monthly_pattern = sales_with_dates.groupby('Month_Name')['Sales Amount'].mean().reindex([
        'January', 'February', 'March', 'April', 'May', 'June',
        'July', 'August', 'September', 'October', 'November', 'December'
    ])
    
    axes[0,1].bar(range(12), monthly_pattern.values/1000, 
                 color=plt.cm.coolwarm(np.linspace(0, 1, 12)))
    axes[0,1].set_title('Seasonal Sales Pattern (Average Monthly)', fontsize=14, fontweight='bold')
    axes[0,1].set_xlabel('Month')
    axes[0,1].set_ylabel('Average Daily Sales (Thousands $)')
    axes[0,1].set_xticks(range(12))
    axes[0,1].set_xticklabels([m[:3] for m in monthly_pattern.index], rotation=45)
    axes[0,1].grid(True, alpha=0.3)
    
    # 3. Quarterly Performance Comparison
    quarterly_sales = sales_with_dates.groupby(['Fiscal Year', 'Fiscal Quarter'])['Sales Amount'].sum().unstack(fill_value=0)
    
    x = np.arange(len(quarterly_sales.index))
    width = 0.2
    quarters = quarterly_sales.columns
    
    for i, quarter in enumerate(quarters):
        axes[1,0].bar(x + i*width, quarterly_sales[quarter]/1000000, width, 
                     label=quarter, alpha=0.8)
    
    axes[1,0].set_title('Quarterly Sales Performance by Fiscal Year', fontsize=14, fontweight='bold')
    axes[1,0].set_xlabel('Fiscal Year')
    axes[1,0].set_ylabel('Sales Amount (Millions $)')
    axes[1,0].set_xticks(x + width * 1.5)
    axes[1,0].set_xticklabels(quarterly_sales.index)
    axes[1,0].legend()
    axes[1,0].grid(True, alpha=0.3)
    
    # 4. Year-over-Year Growth
    yearly_sales = sales_with_dates.groupby('Fiscal Year')['Sales Amount'].sum().reset_index()
    yearly_sales = yearly_sales.sort_values('Fiscal Year')
    yearly_sales['YoY_Growth'] = yearly_sales['Sales Amount'].pct_change() * 100
    
    bars = axes[1,1].bar(yearly_sales['Fiscal Year'][1:], yearly_sales['YoY_Growth'][1:], 
                        color=['green' if x > 0 else 'red' for x in yearly_sales['YoY_Growth'][1:]])
    axes[1,1].set_title('Year-over-Year Sales Growth Rate', fontsize=14, fontweight='bold')
    axes[1,1].set_xlabel('Fiscal Year')
    axes[1,1].set_ylabel('Growth Rate (%)')
    axes[1,1].axhline(y=0, color='black', linestyle='-', alpha=0.3)
    axes[1,1].grid(True, alpha=0.3)
    
    # Add value labels
    for i, bar in enumerate(bars):
        height = bar.get_height()
        axes[1,1].text(bar.get_x() + bar.get_width()/2., height + (1 if height > 0 else -3),
                      f'{height:.1f}%', ha='center', va='bottom' if height > 0 else 'top', 
                      fontweight='bold')
    
    plt.tight_layout()
    plt.savefig('time_series_charts.png', dpi=300, bbox_inches='tight')
    plt.show()

def main():
    """Main function to generate all charts"""
    print("Loading data...")
    data = load_data()
    
    print("Creating Sales Performance Charts...")
    create_sales_performance_charts(data)
    
    print("Creating Geographic Analysis Charts...")
    create_geographic_charts(data)
    
    print("Creating Product Analysis Charts...")
    create_product_charts(data)
    
    print("Creating Customer Analysis Charts...")
    create_customer_charts(data)
    
    print("Creating Channel & Reseller Charts...")
    create_channel_reseller_charts(data)
    
    print("Creating Time Series Analysis Charts...")
    create_time_series_charts(data)
    
    print("\n✅ All charts have been generated and saved as PNG files!")
    print("\nGenerated files:")
    print("- sales_performance_charts.png")
    print("- geographic_analysis_charts.png") 
    print("- product_analysis_charts.png")
    print("- customer_analysis_charts.png")
    print("- channel_reseller_charts.png")
    print("- time_series_charts.png")

if __name__ == "__main__":
    main()