import pandas as pd
import matplotlib.pyplot as plt

# Load the Excel file
file_path = 'financial_data.xlsx'  
expense_df = pd.read_excel(file_path, sheet_name='Expense')
revenue_df = pd.read_excel(file_path, sheet_name='Revenue')
budget_df = pd.read_excel(file_path, sheet_name='Budget')

# Convert dates and extract month
expense_df['Date'] = pd.to_datetime(expense_df['Date'])
revenue_df['Date'] = pd.to_datetime(revenue_df['Date'])
expense_df['Month'] = expense_df['Date'].dt.to_period('M').astype(str)
revenue_df['Month'] = revenue_df['Date'].dt.to_period('M').astype(str)

# Group and summarize
monthly_expense = expense_df.groupby(['Month', 'Department'])['Expense_Amount'].sum().reset_index()
monthly_revenue = revenue_df.groupby('Month')['Revenue_Amount'].sum().reset_index()

# Merge for variance analysis
analysis_df = pd.merge(monthly_expense, budget_df, on=['Month', 'Department'])
analysis_df['Variance'] = analysis_df['Budget'] - analysis_df['Expense_Amount']

# Financial Summary
total_revenue = monthly_revenue['Revenue_Amount'].sum()
total_expense = monthly_expense['Expense_Amount'].sum()
net_profit = total_revenue - total_expense
top_over_budget = analysis_df[analysis_df['Variance'] < 0].sort_values('Variance').head(3)

print("📊 Financial Summary (Jan-Mar 2025):")
print(f"- Total Revenue: ₹{total_revenue:,.0f}")
print(f"- Total Expense: ₹{total_expense:,.0f}")
print(f"- Net Profit: ₹{net_profit:,.0f}")
print("\n🚨 Departments Over Budget:")
for _, row in top_over_budget.iterrows():
    print(f"- {row['Department']} in {row['Month']} exceeded budget by ₹{abs(row['Variance']):,.0f}")

# Plotting Revenue vs Expense
monthly_expense_total = expense_df.groupby('Month')['Expense_Amount'].sum().reset_index()
merged_data = pd.merge(monthly_revenue, monthly_expense_total, on='Month')

plt.figure(figsize=(8, 5))
plt.plot(merged_data['Month'], merged_data['Revenue_Amount'], marker='o', label='Revenue')
plt.plot(merged_data['Month'], merged_data['Expense_Amount'], marker='o', label='Expense')
plt.title('Monthly Revenue vs Expense')
plt.xlabel('Month')
plt.ylabel('Amount (₹)')
plt.legend()
plt.grid(True)
plt.tight_layout()
plt.show()
