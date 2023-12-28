# Henock Zemenfes
# Augmendix Data Home Exercise 

import pandas as pd
import matplotlib.pyplot as plt
import seaborn as sns
from datetime import datetime

# Load the data
file_path = 'Data Analyst - Take Home Test - Data.xlsx'
raw_data = pd.read_excel(file_path, sheet_name='Raw Data')

# Analyze Pipeline Generation
total_pipeline_by_channel = raw_data.groupby('Marketing Channel')['Deal Size'].sum().sort_values(ascending=False)

# Analyze Enterprise and Physician Practice Segments
enterprise_data = raw_data[raw_data['Customer Size'] == 'Enterprise']
physician_practice_data = raw_data[raw_data['Customer Size'] == 'Physician Practice']

enterprise_pipeline_by_channel = enterprise_data.groupby('Marketing Channel')['Deal Size'].sum().sort_values(ascending=False)
physician_practice_pipeline_by_channel = physician_practice_data.groupby('Marketing Channel')['Deal Size'].sum().sort_values(ascending=False)

# Analyze Closed-Won Business
closed_won_data = raw_data[raw_data['Stage'] == 'Contract Signed/Closed Won']
closed_won_by_channel_overall = closed_won_data.groupby('Marketing Channel')['Deal Size'].sum().sort_values(ascending=False)

closed_won_by_channel_enterprise = closed_won_data[closed_won_data['Customer Size'] == 'Enterprise'].groupby('Marketing Channel')['Deal Size'].sum().sort_values(ascending=False)
closed_won_by_channel_physician_practice = closed_won_data[closed_won_data['Customer Size'] == 'Physician Practice'].groupby('Marketing Channel')['Deal Size'].sum().sort_values(ascending=False)

# Calculate Close Rates
close_rates = (closed_won_by_channel_overall / total_pipeline_by_channel).sort_values(ascending=False)

# Trend Analysis
raw_data['Year'] = raw_data['Created Date'].dt.year
pipeline_trend_data = raw_data.groupby(['Marketing Channel', 'Year'])['Deal Size'].sum().reset_index()

# Plotting the trends
plt.figure(figsize=(15, 8))
sns.lineplot(data=pipeline_trend_data, x='Year', y='Deal Size', hue='Marketing Channel', marker='o')
plt.title('Pipeline Generation Trend by Marketing Channel (2022-2023)')
plt.xlabel('Year')
plt.ylabel('Total Pipeline Generated ($)')
plt.legend(title='Marketing Channel', bbox_to_anchor=(1.05, 1), loc='upper left')
plt.tight_layout()
plt.show()

# Additional Insights
# Sales Rep Performance
top_sales_reps = raw_data.groupby('Opportunity Owner')['Deal Size'].sum().sort_values(ascending=False).head(5)

# Deal Size Distribution
deal_size_distribution = raw_data['Deal Size'].describe()

# Customer Size Distribution
customer_size_distribution = raw_data['Customer Size'].value_counts()
