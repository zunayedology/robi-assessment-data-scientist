# Importing necessary libraries
import pandas as pd

# Loading the predictions from Task 2
predictions = pd.read_excel('sales_predictions.xlsx', index_col=0)

# Loading the outlet information to understand the serving schedule
# 'sheet_name' specifies the sheet from which to read data
outlet_info = pd.read_excel('Worksheet in Assessment question (006).xlsx',
                            sheet_name='outlet info')

total_amounts = {}

# Calculating the total balance needed for each outlet
for outlet, sales in predictions.iterrows():
    # Get the serving schedule for the current outlet
    serving_days = outlet_info[outlet_info['outlet'] == outlet].iloc[0][2:9]

    # Identify which days the outlet is served based on the serving schedule
    next_visit_days = serving_days[serving_days == 1].index.tolist()

    # Calculate the balance needed until the next visit
    total_amount = 0
    for day in ['2024-02-01', '2024-02-02', '2024-02-03']:
        total_amount += sales[day]

        # Accumulate the sales amount for each day
        if day in next_visit_days:
            break

    total_amounts[outlet] = total_amount

# Convert the total amounts dictionary to a DataFrame for output
# 'from_dict' creates a DataFrame from the dictionary with outlet names as index and a column for total balance
total_amounts_df = pd.DataFrame.from_dict(total_amounts,
                                          orient='index',
                                          columns=['Total Balance Needed'])

# Save the total amounts DataFrame to an Excel file
total_amounts_df.to_excel('total_balance_needed.xlsx')
