import pandas as pd
import os
import re
from datetime import datetime
import matplotlib.pyplot as plt
import matplotlib
matplotlib.use('Agg')  # Use non-interactive backend

# ---- CONFIG ----
input_folder = "processed_output"
output_file = "processed_output/analysis_report.xlsx"

# Category keywords for automatic categorization
CATEGORY_KEYWORDS = {
    'Food & Dining': [
        'GRAB-EC', 'KEN\'S GROCER', 'FINEXUS CAFE', 'McDonalds', 'McDonald',
        'RESTAURANT', 'CAFE', 'COFFEE', 'FOOD', 'NASI LEMAK', 'BISTRO',
        'TAO BIN', 'LIKEMECOFFEE', 'ZUS COFFEE', 'SARA THAI', 'LOK\'AI',
        'DUITNOWQR', 'QR PAYMENT', 'GORENG', 'NASI'
    ],
    'Transport': [
        'MRT', 'UBER', 'TAXI', 'PARKING', 'TOLL', 'PETROL', 'GAS', 'PUTRAJAYA SENTRAL'
    ],
    'Bills & Utilities': [
        'U MOBILE', 'MOBILE', 'PHONE', 'UTILITY', 'BILL', 'ELECTRIC', 'WATER',
        'INTERNET', 'PAYBILL'
    ],
    'Entertainment': [
        'STEAMGAMES', 'SPOTIFY', 'NETFLIX', 'CINEMA', 'MOVIE', 'ENTERTAINMENT',
        'GAME', 'GAMES'
    ],
    'Shopping': [
        'AEON', 'SHOPPING', 'MALL', 'STORE', 'RETAIL', 'SUPERMARKET'
    ],
    'Transfers': [
        'DUITNOW', 'TRANSFER', 'FUND TRANSFER', 'FUTU MALAYSIA', 'MOOMOO'
    ],
    'Savings & Interest': [
        'PROFIT EARNED', 'INTEREST', 'SAVINGS ACCOUNT'
    ]
}

def extract_month_year_from_filename(filename):
    """Extract month and year from filename like '1000073282_202501.csv'."""
    # Pattern: accountnumber_YYYYMM.csv
    match = re.search(r'_(\d{6})\.csv$', filename)
    if match:
        period_str = match.group(1)
        year = int(period_str[:4])
        month = int(period_str[4:6])
        return year, month
    return None, None

def categorize_transaction(description):
    """Categorize a transaction based on its description."""
    if pd.isna(description):
        return 'Other'
    
    desc_upper = str(description).upper()
    
    # Check each category
    for category, keywords in CATEGORY_KEYWORDS.items():
        for keyword in keywords:
            if keyword.upper() in desc_upper:
                return category
    
    return 'Other'

def load_all_csv_files(folder_path):
    """Load and combine all CSV files from the folder."""
    all_dataframes = []
    
    if not os.path.exists(folder_path):
        print(f"ERROR: Folder '{folder_path}' does not exist!")
        return None
    
    csv_files = [f for f in os.listdir(folder_path) if f.endswith('.csv')]
    
    if len(csv_files) == 0:
        print(f"ERROR: No CSV files found in '{folder_path}'!")
        return None
    
    print(f"Found {len(csv_files)} CSV file(s)")
    
    for csv_file in sorted(csv_files):
        csv_path = os.path.join(folder_path, csv_file)
        print(f"Loading: {csv_file}")
        
        try:
            df = pd.read_csv(csv_path)
            
            # Extract month and year from filename
            year, month = extract_month_year_from_filename(csv_file)
            if year and month:
                df['Year'] = year
                df['Month'] = month
                df['Month_Year'] = pd.to_datetime(f"{year}-{month:02d}-01")
            else:
                print(f"Warning: Could not extract date from filename {csv_file}")
                df['Year'] = None
                df['Month'] = None
                df['Month_Year'] = None
            
            # Add source file column
            df['Source_File'] = csv_file
            
            all_dataframes.append(df)
        except Exception as e:
            print(f"Error loading {csv_file}: {e}")
    
    if len(all_dataframes) == 0:
        return None
    
    # Combine all dataframes
    combined_df = pd.concat(all_dataframes, ignore_index=True)
    print(f"\nTotal transactions loaded: {len(combined_df)}")
    
    return combined_df

def process_transactions(df):
    """Process and enrich the transaction data."""
    if df is None or len(df) == 0:
        return df
    
    # Filter out opening/closing balance entries
    df = df[~df['Description'].astype(str).str.contains('Opening balance|Closing balance', case=False, na=False)].copy()
    
    if len(df) == 0:
        print("Warning: No transactions found after filtering opening/closing balances")
        return df
    
    # Parse dates
    df['Date_Parsed'] = pd.to_datetime(df['Date'], format='%d %b %y', errors='coerce', dayfirst=True)
    
    # Add category
    df['Category'] = df['Description'].apply(categorize_transaction)
    
    # Ensure numeric columns
    df['Withdrawal'] = pd.to_numeric(df['Withdrawal'], errors='coerce').fillna(0)
    df['Deposit'] = pd.to_numeric(df['Deposit'], errors='coerce').fillna(0)
    df['Balance'] = pd.to_numeric(df['Balance'], errors='coerce').fillna(0)
    
    # Calculate transaction amount (positive for deposits, negative for withdrawals)
    df['Amount'] = df['Deposit'] - df['Withdrawal']
    
    # Sort by date
    df = df.sort_values('Date_Parsed', na_position='last')
    df = df.reset_index(drop=True)
    
    return df

def generate_monthly_summary(df):
    """Generate monthly summary statistics."""
    monthly_data = []
    
    if df is None or len(df) == 0:
        return pd.DataFrame()
    
    # Filter out rows with missing Year or Month
    df_valid = df[(df['Year'].notna()) & (df['Month'].notna())].copy()
    
    if len(df_valid) == 0:
        return pd.DataFrame()
    
    for (year, month), group in df_valid.groupby(['Year', 'Month']):
        month_year = pd.to_datetime(f"{int(year)}-{int(month):02d}-01")
        
        total_withdrawals = group['Withdrawal'].sum()
        total_deposits = group['Deposit'].sum()
        net_cash_flow = total_deposits - total_withdrawals
        num_transactions = len(group)
        
        # Calculate average daily spending (only for days with withdrawals)
        days_with_spending = group[group['Withdrawal'] > 0]['Date_Parsed'].dt.date.nunique()
        avg_daily_spending = total_withdrawals / days_with_spending if days_with_spending > 0 else 0
        
        # Category breakdown
        category_spending = group.groupby('Category')['Withdrawal'].sum().to_dict()
        
        row_data = {
            'Year': year,
            'Month': month,
            'Month_Name': month_year.strftime('%B %Y'),
            'Total_Withdrawals': total_withdrawals,
            'Total_Deposits': total_deposits,
            'Net_Cash_Flow': net_cash_flow,
            'Number_of_Transactions': num_transactions,
            'Days_with_Spending': days_with_spending,
            'Avg_Daily_Spending': avg_daily_spending
        }
        
        # Add category columns (sanitize category names for column names)
        for cat in CATEGORY_KEYWORDS.keys():
            col_name = f'Category_{cat.replace(" & ", "_").replace(" ", "_")}'
            row_data[col_name] = category_spending.get(cat, 0)
        
        monthly_data.append(row_data)
    
    summary_df = pd.DataFrame(monthly_data)
    summary_df = summary_df.sort_values(['Year', 'Month'])
    
    return summary_df

def generate_category_summary(df):
    """Generate category-wise spending summary."""
    category_data = []
    
    if df is None or len(df) == 0:
        return pd.DataFrame()
    
    for category in sorted(df['Category'].unique()):
        cat_df = df[df['Category'] == category]
        total_spending = cat_df['Withdrawal'].sum()
        num_transactions = len(cat_df)
        avg_transaction = total_spending / num_transactions if num_transactions > 0 else 0
        
        # Monthly breakdown
        monthly_spending = cat_df.groupby(['Year', 'Month'])['Withdrawal'].sum()
        
        total_all_spending = abs(df['Withdrawal'].sum())
        percentage = (abs(total_spending) / total_all_spending * 100) if total_all_spending > 0 else 0
        
        category_data.append({
            'Category': category,
            'Total_Spending': abs(total_spending),  # Store as positive for readability
            'Number_of_Transactions': num_transactions,
            'Avg_Transaction_Amount': abs(avg_transaction),
            'Percentage_of_Total': percentage
        })
    
    category_df = pd.DataFrame(category_data)
    category_df = category_df.sort_values('Total_Spending', ascending=False)
    
    return category_df

def generate_top_merchants(df, top_n=10):
    """Generate top merchants analysis."""
    # Extract merchant name from description (first part before dash or common patterns)
    def extract_merchant(desc):
        if pd.isna(desc):
            return 'Unknown'
        desc_str = str(desc)
        # Try to extract merchant name (usually before first dash or before payment method)
        parts = desc_str.split(' - ')
        if len(parts) > 0:
            merchant = parts[0].strip()
            # Clean up common patterns
            merchant = re.sub(r'\s+', ' ', merchant)
            return merchant
        return desc_str[:50]  # Truncate if too long
    
    df['Merchant'] = df['Description'].apply(extract_merchant)
    
    merchant_stats = df.groupby('Merchant').agg({
        'Withdrawal': ['sum', 'count'],
        'Deposit': 'sum'
    }).reset_index()
    
    merchant_stats.columns = ['Merchant', 'Total_Spending', 'Transaction_Count', 'Total_Deposits']
    # Convert spending to positive for sorting and display
    merchant_stats['Total_Spending'] = merchant_stats['Total_Spending'].abs()
    merchant_stats = merchant_stats.sort_values('Total_Spending', ascending=False)
    
    # Top by spending
    top_by_spending = merchant_stats.head(top_n).copy()
    top_by_spending['Rank'] = range(1, len(top_by_spending) + 1)
    
    # Top by frequency
    top_by_frequency = merchant_stats.sort_values('Transaction_Count', ascending=False).head(top_n).copy()
    top_by_frequency['Rank'] = range(1, len(top_by_frequency) + 1)
    
    return top_by_spending, top_by_frequency

def create_visualizations(df, monthly_summary, category_summary, output_dir='processed_output'):
    """Create visualization charts."""
    charts = {}
    
    # 1. Monthly Spending Trend
    plt.figure(figsize=(12, 6))
    monthly_summary_sorted = monthly_summary.sort_values(['Year', 'Month'])
    plt.plot(monthly_summary_sorted['Month_Name'], monthly_summary_sorted['Total_Withdrawals'], 
             marker='o', linewidth=2, markersize=8)
    plt.title('Monthly Spending Trend', fontsize=14, fontweight='bold')
    plt.xlabel('Month', fontsize=12)
    plt.ylabel('Total Spending (RM)', fontsize=12)
    plt.xticks(rotation=45, ha='right')
    plt.grid(True, alpha=0.3)
    plt.tight_layout()
    chart1_path = os.path.join(output_dir, 'chart_monthly_trend.png')
    plt.savefig(chart1_path, dpi=150, bbox_inches='tight')
    plt.close()
    charts['Monthly_Trend'] = chart1_path
    
    # 2. Category Breakdown (Pie Chart)
    plt.figure(figsize=(10, 8))
    category_summary_filtered = category_summary[category_summary['Total_Spending'] > 0]
    plt.pie(category_summary_filtered['Total_Spending'], 
            labels=category_summary_filtered['Category'],
            autopct='%1.1f%%', startangle=90)
    plt.title('Spending by Category', fontsize=14, fontweight='bold')
    plt.tight_layout()
    chart2_path = os.path.join(output_dir, 'chart_category_pie.png')
    plt.savefig(chart2_path, dpi=150, bbox_inches='tight')
    plt.close()
    charts['Category_Pie'] = chart2_path
    
    # 3. Monthly Category Comparison (Stacked Bar)
    plt.figure(figsize=(14, 8))
    monthly_summary_sorted = monthly_summary.sort_values(['Year', 'Month'])
    categories = sorted(df['Category'].unique())
    category_colors = plt.cm.Set3(range(len(categories)))
    
    bottom = None
    for i, category in enumerate(categories):
        values = []
        for _, row in monthly_summary_sorted.iterrows():
            col_name = f'Category_{category.replace(" & ", "_").replace(" ", "_")}'
            values.append(row.get(col_name, 0))
        
        if bottom is None:
            plt.bar(monthly_summary_sorted['Month_Name'], values, label=category, color=category_colors[i])
            bottom = values
        else:
            plt.bar(monthly_summary_sorted['Month_Name'], values, bottom=bottom, 
                   label=category, color=category_colors[i])
            bottom = [b + v for b, v in zip(bottom, values)]
    
    plt.title('Monthly Spending by Category', fontsize=14, fontweight='bold')
    plt.xlabel('Month', fontsize=12)
    plt.ylabel('Total Spending (RM)', fontsize=12)
    plt.xticks(rotation=45, ha='right')
    plt.legend(bbox_to_anchor=(1.05, 1), loc='upper left')
    plt.grid(True, alpha=0.3, axis='y')
    plt.tight_layout()
    chart3_path = os.path.join(output_dir, 'chart_monthly_category.png')
    plt.savefig(chart3_path, dpi=150, bbox_inches='tight')
    plt.close()
    charts['Monthly_Category'] = chart3_path
    
    # 4. Top Merchants (Bar Chart)
    plt.figure(figsize=(12, 8))
    top_merchants, _ = generate_top_merchants(df, top_n=10)
    plt.barh(range(len(top_merchants)), top_merchants['Total_Spending'], 
             color=plt.cm.viridis(range(len(top_merchants))))
    plt.yticks(range(len(top_merchants)), top_merchants['Merchant'])
    plt.xlabel('Total Spending (RM)', fontsize=12)
    plt.title('Top 10 Merchants by Spending', fontsize=14, fontweight='bold')
    plt.gca().invert_yaxis()
    plt.grid(True, alpha=0.3, axis='x')
    plt.tight_layout()
    chart4_path = os.path.join(output_dir, 'chart_top_merchants.png')
    plt.savefig(chart4_path, dpi=150, bbox_inches='tight')
    plt.close()
    charts['Top_Merchants'] = chart4_path
    
    return charts

def export_to_excel(df, monthly_summary, category_summary, top_merchants_spending, 
                    top_merchants_frequency, charts, output_path):
    """Export all data and charts to Excel file."""
    print(f"\nExporting to Excel: {output_path}")
    
    with pd.ExcelWriter(output_path, engine='openpyxl') as writer:
        # Sheet 1: All Transactions
        df_export = df[['Date', 'Description', 'Category', 'Withdrawal', 'Deposit', 
                        'Balance', 'Year', 'Month', 'Source_File']].copy()
        df_export.to_excel(writer, sheet_name='All Transactions', index=False)
        
        # Sheet 2: Monthly Summary
        monthly_summary.to_excel(writer, sheet_name='Monthly Summary', index=False)
        
        # Sheet 3: Category Summary
        category_summary.to_excel(writer, sheet_name='Category Summary', index=False)
        
        # Sheet 4: Top Merchants
        top_merchants_spending.to_excel(writer, sheet_name='Top Merchants (Spending)', index=False)
        top_merchants_frequency.to_excel(writer, sheet_name='Top Merchants (Frequency)', index=False)
        
        # Sheet 5: Charts (add chart references)
        charts_df = pd.DataFrame({
            'Chart Name': list(charts.keys()),
            'Chart File': list(charts.values())
        })
        charts_df.to_excel(writer, sheet_name='Charts', index=False)
    
    print(f"Excel file created successfully!")

def main():
    """Main analysis function."""
    print("="*60)
    print("Bank Statement Analyzer")
    print("="*60)
    
    # Load all CSV files
    print("\n[1/6] Loading CSV files...")
    df = load_all_csv_files(input_folder)
    if df is None:
        return
    
    # Process transactions
    print("\n[2/6] Processing transactions...")
    df = process_transactions(df)
    print(f"Processed {len(df)} transactions")
    print(f"Date range: {df['Date_Parsed'].min()} to {df['Date_Parsed'].max()}")
    
    # Generate monthly summary
    print("\n[3/6] Generating monthly summaries...")
    monthly_summary = generate_monthly_summary(df)
    print(f"Generated summary for {len(monthly_summary)} months")
    
    # Generate category summary
    print("\n[4/6] Generating category summaries...")
    category_summary = generate_category_summary(df)
    print(f"Found {len(category_summary)} categories")
    
    # Generate top merchants
    print("\n[5/6] Analyzing top merchants...")
    top_merchants_spending, top_merchants_frequency = generate_top_merchants(df, top_n=10)
    print(f"Top merchants identified")
    
    # Create visualizations
    print("\n[6/6] Creating visualizations...")
    charts = create_visualizations(df, monthly_summary, category_summary, input_folder)
    print(f"Created {len(charts)} charts")
    
    # Export to Excel
    print("\nExporting to Excel...")
    export_to_excel(df, monthly_summary, category_summary, top_merchants_spending,
                   top_merchants_frequency, charts, output_file)
    
    # Print summary statistics
    print("\n" + "="*60)
    print("ANALYSIS SUMMARY")
    print("="*60)
    total_withdrawals = abs(df['Withdrawal'].sum())
    total_deposits = df['Deposit'].sum()
    
    print(f"Total Transactions: {len(df):,}")
    print(f"Total Withdrawals: RM {total_withdrawals:,.2f}")
    print(f"Total Deposits: RM {total_deposits:,.2f}")
    print(f"Net Cash Flow: RM {total_deposits - total_withdrawals:,.2f}")
    print(f"\nTop Categories:")
    for _, row in category_summary.head(5).iterrows():
        print(f"  {row['Category']}: RM {row['Total_Spending']:,.2f} ({row['Percentage_of_Total']:.1f}%)")
    print(f"\nOutput saved to: {output_file}")
    print("="*60)

if __name__ == "__main__":
    main()

