import pandas as pd
import numpy as np
import sys
import re

# Load the Qualtrics CSV into a DataFrame and specify 'IPAddress' as a string
qualtrics_df = pd.read_csv('Qualtrics3.csv', delimiter=',', dtype={'IPAddress': str})

# Convert the 'RecordedDate' and 'StartDate' columns in Qualtrics to datetime
qualtrics_df['RecordedDate'] = pd.to_datetime(qualtrics_df['RecordedDate'], format='%m/%d/%y %H:%M')
qualtrics_df['StartDate'] = pd.to_datetime(qualtrics_df['StartDate'], format='%m/%d/%y %H:%M')

# Truncate seconds and keep up to minutes
qualtrics_df['RecordedDate'] = qualtrics_df['RecordedDate'].dt.floor('T')

# Add a 'Matched/Not Matched' column and 'Repeated User' column
qualtrics_df['Matched/Not Matched'] = 'Not Matched'
qualtrics_df['Repeated User'] = 'No'

# Load the filtered_data.csv into a DataFrame
filtered_df = pd.read_csv('filtered_data.csv')

# Add two empty columns after 'Repeated User'
qualtrics_df[' '], qualtrics_df[' '] = np.nan, np.nan

# Initialize new columns with NaN values
new_columns = ['idVisit', 'visitIp', 'serverTimePretty', 'visitDurationPretty', 'referrerName', 'referrerKeyword']
for col in new_columns:
    qualtrics_df[col] = np.nan

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
                qualtrics_df.at[index, 'Matched/Not Matched'] = 'Matched'
                for new_col in new_columns:
                    qualtrics_df.at[index, new_col] = matching_rows.iloc[0][new_col]

# Find repeated users based on IP
# Count the number of non-null 'serverTimePretty (actionDetails X)' for each IP
count_by_ip = filtered_df.filter(like='serverTimePretty (actionDetails').apply(lambda x: x.count(), axis=1)
# Filter out IPs with count >= 2
repeated_users = filtered_df.loc[count_by_ip >= 2, 'visitIp'].unique()

# Convert 'IPAddress' to string type in case it's not
qualtrics_df['IPAddress'] = qualtrics_df['IPAddress'].astype(str)

# Update 'Repeated User' column in qualtrics_df
for ip in repeated_users:
    pattern = r'^' + re.escape(ip.split('.')[0]) + r'\.' + re.escape(ip.split('.')[1])
    qualtrics_df.loc[qualtrics_df['IPAddress'].str.match(pattern, case=False), 'Repeated User'] = 'Yes'

# Sort rows by 'StartDate'
qualtrics_df.sort_values(by='StartDate', inplace=True)

# Save the matching rows to a new CSV file
qualtrics_df.to_csv('Qualtrics_matching.csv', index=False)

# Terminate the program
sys.exit(0)
