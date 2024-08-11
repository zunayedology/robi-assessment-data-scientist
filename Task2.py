# Importing necessary libraries
import pandas as pd
from statsmodels.tsa.arima.model import ARIMA  # ARIMA is an excellent model for time-series forecasting

# Loading sales information from an Excel file
# The 'sheet_name' parameter specifies the sheet from which to read data
sales_info = pd.read_excel('Worksheet in Assessment question (006).xlsx',
                           sheet_name='sales info')

# Converting the 'date' column to datetime format
# 'pd.to_datetime' converts the date strings to datetime objects
sales_info['date'] = pd.to_datetime(sales_info['date'],
                                    format='%d-%b-%y')

# Pivoting the sales data to have dates as rows and outlets as columns
# 'pivot' reshapes the DataFrame, and 'fillna(0)' replaces any NaN values with 0
sales_data = sales_info.pivot(index='date', columns='outlet', values='sales BDT').fillna(0)

predictions = {}

# Training the model for each outlet and making predictions
for outlet in sales_data.columns:
    y = sales_data[outlet].values

    # Creating and fitting the model
    model = ARIMA(y, order=(5, 1, 0))
    model_fit = model.fit()

    # Predict sales for February 1, 2, and 3
    pred_dates = [len(sales_data) + i for i in range(1, 4)]  # Generating forecast steps
    forecast = model_fit.forecast(steps=3)  # Forecasting the next 3 time steps
    predictions[outlet] = forecast  # Storing the forecast in the dictionary

# Converting the dictionary to a DataFrame
# 'DataFrame' creates a DataFrame with dates as index and outlets as columns, and '.T' transposes it
prediction_df = pd.DataFrame(predictions,
                             index=['2024-02-01',
                                    '2024-02-02',
                                    '2024-02-03']).T

# Save the predictions DataFrame to an Excel file
prediction_df.to_excel('sales_predictions.xlsx')
