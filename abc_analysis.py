import pandas as pd

# Load data from Excel
file_path = 'C:\\Users\\Frank\\Desktop\\stock_data.xlsx'  # Replace with your actual file path
df = pd.read_excel(file_path)

# Since the Value column represents the total value for each code, we can directly use it.
df['Total Value'] = df['Value']

# Sort items by total value in descending order
df = df.sort_values(by='Total Value', ascending=False)

# Calculate cumulative sum of total value
df['Cumulative Sum'] = df['Total Value'].cumsum()

# Calculate cumulative percentage
total_value_sum = df['Total Value'].sum()
df['Cumulative Percentage'] = df['Cumulative Sum'] / total_value_sum * 100

# Assign ABC category
def assign_abc_category(row):
    if row['Cumulative Percentage'] <= 80:
        return 'A'
    elif row['Cumulative Percentage'] <= 95:
        return 'B'
    else:
        return 'C'

df['ABC Category'] = df.apply(assign_abc_category, axis=1)

# Display the DataFrame
print(df[['Stock Code', 'Stock Description', 'Quantity', 'Value', 'Total Value', 'Cumulative Percentage', 'ABC Category']])

# Optional: Plotting with Plotly
import plotly.express as px

fig = px.bar(df, x='Stock Code', y='Total Value', color='ABC Category',
             title='ABC Analysis of Stock Items',
             labels={'Total Value': 'Total Value ($)', 'Stock Code': 'Stock Code'},
             height=600)
fig.show()
