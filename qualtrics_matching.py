import pandas as pd
import numpy as np
import re
import sys
import xlsxwriter

# Load the Qualtrics CSV into a DataFrame and specify 'IPAddress' as a string
qualtrics_df = pd.read_csv('Qualtrics1.csv', delimiter=',', dtype={'IPAddress': str})

# Convert the 'RecordedDate' and 'StartDate' columns in Qualtrics to datetime
qualtrics_df['RecordedDate'] = pd.to_datetime(qualtrics_df['RecordedDate'], format='%m/%d/%y %H:%M')
qualtrics_df['StartDate'] = pd.to_datetime(qualtrics_df['StartDate'], format='%m/%d/%y %H:%M')

# Truncate seconds and keep up to minutes
qualtrics_df['RecordedDate'] = qualtrics_df['RecordedDate'].dt.floor('T')

# Add empty columns and new columns for matching
for _ in range(4):  # Adding 4 empty columns
    qualtrics_df['                 '] = np.nan

# Add 'IP and TimeStamp Matched/Not Matched' column and 'Repeated User' column
qualtrics_df['TimeStamp Matched/Not Matched'] = 'Not Matched'
qualtrics_df['IP Matched/Not Matched'] = 'Not Matched'
qualtrics_df['Repeated User'] = 'No'

for _ in range(4):  # Adding 4 more empty columns
    qualtrics_df['                '] = np.nan

# Load the filtered_data.csv into a DataFrame
filtered_df = pd.read_csv('filtered_data.csv')

# Get unique IPs from filtered_data.csv
filtered_ips = filtered_df['visitIp'].unique()

# Initialize new columns with NaN values
new_columns = ['idVisit', 'visitIp', 'serverTimePretty', 'visitDurationPretty', 'referrerName', 'referrerKeyword']
for col in new_columns:
    qualtrics_df[col] = np.nan

for _ in range(4):  # Adding 4 more empty columns
    qualtrics_df['                 '] = np.nan

# Loop through each 'serverTimePretty (actionDetails X)' column in filtered_data.csv
for col in filtered_df.columns:
    if 'serverTimePretty (actionDetails' in col:
        # Convert the column to datetime
        filtered_df[col] = pd.to_datetime(filtered_df[col], format='%b %d, %Y %H:%M:%S')
        filtered_df[col] = filtered_df[col].dt.floor('T')

        # Update the 'Matched' column and new columns in qualtrics_df for matching rows
        for index, row in qualtrics_df.iterrows():
            matching_rows = filtered_df[filtered_df[col] == row['RecordedDate']]
            if not matching_rows.empty:
                qualtrics_df.at[index, 'TimeStamp Matched/Not Matched'] = 'Matched'
                for new_col in new_columns:
                    qualtrics_df.at[index, new_col] = matching_rows.iloc[0][new_col]

# Check if each truncated IP in qualtrics_df is a prefix in filtered_data.csv
for index, row in qualtrics_df.iterrows():
    if any(ip.startswith(row['IPAddress']) for ip in filtered_ips):
        qualtrics_df.at[index, 'IP Matched/Not Matched'] = 'Matched'

# Find repeated users based on IP
count_by_ip = filtered_df.filter(like='serverTimePretty (actionDetails').apply(lambda x: x.count(), axis=1)
repeated_users = filtered_df.loc[count_by_ip >= 2, 'visitIp'].unique()

qualtrics_df['IPAddress'] = qualtrics_df['IPAddress'].astype(str)

for ip in repeated_users:
    pattern = r'^' + re.escape(ip.split('.')[0]) + r'\.' + re.escape(ip.split('.')[1])
    qualtrics_df.loc[qualtrics_df['IPAddress'].str.match(pattern, case=False), 'Repeated User'] = 'Yes'

# To determine combined matching for timestamp and IP
qualtrics_df['IP & TimeStamp Combined'] = np.where(
    (qualtrics_df['TimeStamp Matched/Not Matched'] == 'Matched') &
    (qualtrics_df['IP Matched/Not Matched'] == 'Matched'),
    'Both Matched', 'Not Both Matched'
)

# Sort rows by 'StartDate'
qualtrics_df.sort_values(by='StartDate', inplace=True)

# Replace NaN values with an empty string
qualtrics_df = qualtrics_df.fillna('')

# Open a new Excel file and add a worksheet.
workbook = xlsxwriter.Workbook('Qualtrics_Validation_Facebook_09-11To09-22.xlsx', {'nan_inf_to_errors': True})
worksheet = workbook.add_worksheet()

# Define Excel formats
header_format = workbook.add_format({
    'bold': True,
    'fg_color': '#D7E4BC',
    'border': 1,
    'text_wrap': True,
    'valign': 'vcenter',
    'align': 'center'
})
cell_format = workbook.add_format({
    'border': 1
})

# Write DataFrame data to the Excel sheet
for row_num, (index, row_data) in enumerate(qualtrics_df.iterrows()):
    for col_num, cell_data in enumerate(row_data):
        worksheet.write(row_num + 1, col_num, cell_data, cell_format)

# Apply column headers and formatting
for col_num, column_name in enumerate(qualtrics_df.columns):
    worksheet.write(0, col_num, column_name, header_format)

# Adjust column width to fit data
for i, column in enumerate(qualtrics_df):
    max_len = max(
        qualtrics_df[column].astype(str).apply(len).max(),  # len of largest item
        len(str(column))  # len of column name/header
    )
    worksheet.set_column(i, i, max_len)  # set column width

# Define pie chart colors
chart_colors = ['#00FF00', '#FF0000']

# Pie chart data
categories = {
    'TimeStamp Matched/Not Matched': qualtrics_df['TimeStamp Matched/Not Matched'].value_counts().reindex(['Matched', 'Not Matched']).fillna(0),
    'IP Matched/Not Matched': qualtrics_df['IP Matched/Not Matched'].value_counts().reindex(['Matched', 'Not Matched']).fillna(0),
    'IP & TimeStamp Combined': qualtrics_df['IP & TimeStamp Combined'].value_counts().reindex(['Both Matched', 'Not Both Matched']).fillna(0)
}

# Create pie charts using xlsxwriter
start_col = len(qualtrics_df.columns) + 2
for idx, (title, values) in enumerate(categories.items()):
    end_row = len(values) + 1

    data_start_col = start_col + 2*idx
    data_end_col = data_start_col + 1

    # Insert data for chart
    for row_num, (label, value) in enumerate(values.items(), start=1):
        worksheet.write(row_num, data_start_col, label)
        worksheet.write(row_num, data_end_col, value)

    # Create a new pie chart object.
    pie_chart = workbook.add_chart({'type': 'pie'})

    # Configure the chart, with a title defined by the category.
    pie_chart.add_series({
        'name': title,
        'categories': ['Sheet1', 1, data_start_col, end_row, data_start_col],
        'values': ['Sheet1', 1, data_end_col, end_row, data_end_col],
        'points': [{'fill': {'color': chart_colors[i]}} for i in range(len(values))],
        'data_labels': {'percentage': True},
    })
    pie_chart.set_title({'name': title})

    # Adjusted pie chart positioning
    chart_position_row = 0
    chart_position_col = data_end_col + 2
    worksheet.insert_chart(chart_position_row, chart_position_col, pie_chart)

# Close the Pandas Excel writer and output the Excel file.
workbook.close()
sys.exit(0)
