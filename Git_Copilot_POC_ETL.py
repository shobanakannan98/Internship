import pandas as pd

# Load the Excel file
file_path = r"C:\Users\perumalm\Desktop\GIT\MySQL server\GitHub_Copliot_demo\Excelfiles\Git_Copilot_POC_B1.xlsx"
df = pd.read_excel(file_path, sheet_name=None)

# Extract the sheets
dim_customer = df['DimCustomer']
dim_product = df['DimProduct']
fact_finance = df['FactFinance']

# Print the columns of the FactFinance DataFrame for inspection
print("Columns in FactFinance:", fact_finance.columns)


# Transform the FactFinance table
fact_finance_actual = fact_finance[['Order Quantity, actual']].copy()
fact_finance_actual['Scenario'] = 'Actual'
fact_finance_actual.rename(columns={'Order Quantity, actual': 'OrderQuantity'}, inplace=True)

fact_finance_budget = fact_finance[['Order Quantity, budget']].copy()
fact_finance_budget['Scenario'] = 'Budget'
fact_finance_budget.rename(columns={'Order Quantity, budget': 'OrderQuantity'}, inplace=True)

# Combine the transformed data
fact_finance_transformed = pd.concat([fact_finance_actual, fact_finance_budget], ignore_index=True)
# Clean the 'Month' column to ensure all values are in the correct format
fact_finance['Month'] = fact_finance['Month'].apply(lambda x: str(x).zfill(6) if pd.notnull(x) else x)

# Convert the 'Month' column to a date column in the FactFinance sheet
fact_finance['Date'] = pd.to_datetime(fact_finance['Month'].astype(str), format='%Y%m', errors='coerce').dt.strftime('%d.%m.%Y')

# Drop the old 'Month' column
fact_finance.drop(columns=['Month'], inplace=True)

# Save the transformed data to a new Excel file
output_file_path = r"C:\Users\perumalm\Desktop\GIT\MySQL server\GitHub_Copliot_demo\Excelfiles\Git_Copilot_POC_Transformed_B1.xlsx"
with pd.ExcelWriter(output_file_path) as writer:
    dim_customer.to_excel(writer, sheet_name='DimCustomer', index=False)
    dim_product.to_excel(writer, sheet_name='DimProduct', index=False)
    fact_finance_transformed.to_excel(writer, sheet_name='FactFinance', index=False)
    fact_finance.to_excel(writer, sheet_name='FactFinance', index=False)

print(f"Transformed data has been saved to {output_file_path}")
