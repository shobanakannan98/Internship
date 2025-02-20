import pandas as pd

# Load the Excel file
file_path = r'C:\Users\perumalm\Desktop\GIT\MySQL server\GitHub_Copliot_demo\Excelfiles\Git_Copilot_POC_copy_B1.xlsx'
xls = pd.ExcelFile(file_path)

# Read the FactFinance sheet
df_fact_finance = pd.read_excel(xls, 'FactFinance')

# Ensure the column names are stripped of spaces (in case of unexpected white spaces)
df_fact_finance.columns = df_fact_finance.columns.str.strip()

# Transform the columns in FactFinance
df_actual = df_fact_finance.copy()
df_actual['Scenario'] = 'Actual'

df_actual['Sales Amount'] = df_actual['Sales Amount, actual']
df_actual['Total Product Cost'] = df_actual['Total Product Cost, actual']
df_actual['Fixed Costs'] = df_actual['Fixed Costs, actual']

df_budget = df_fact_finance.copy()
df_budget['Scenario'] = 'Budget'

df_budget['Sales Amount'] = df_budget['Sales Amount, budget']
df_budget['Total Product Cost'] = df_budget['Total Product Cost, budget']
df_budget['Fixed Costs'] = df_budget['Fixed Costs, budget']

# Concatenate the transformed data
df_transformed = pd.concat([df_actual, df_budget], ignore_index=True)

# Drop the original actual/budget columns
df_transformed.drop(columns=[
    
    'Sales Amount, actual', 'Sales Amount, budget',
    'Total Product Cost, actual', 'Total Product Cost, budget',
    'Fixed Costs, actual', 'Fixed Costs, budget'
], inplace=True)

# Save the transformed data back to the Excel file
with pd.ExcelWriter(file_path, engine='openpyxl', mode='a', if_sheet_exists='replace') as writer:
    df_transformed.to_excel(writer, sheet_name='FactFinance', index=False)

print("Transformation is successfully saved.")


