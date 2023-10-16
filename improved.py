import pandas as pd
import numpy as np
import re
import xlsxwriter
from pathlib import Path
import warnings
warnings.filterwarnings("ignore")


class DataProcessor:
    def __init__(self, qualtrics_filename, matomo_data_filename, output_filename):
        self.qualtrics_file = Path(qualtrics_filename)
        self.matomo_file = Path(matomo_data_filename)
        self.output_file = Path(output_filename)

    def process_matomo(self):

        # Read the CSV file into a DataFrame with the correct delimiter and encoding
        df = pd.read_csv(self.matomo_file, delimiter=',', encoding='utf-16')

        # Create a list of specific columns you want to keep
        specific_columns = ['idVisit', 'visitIp', 'serverTimePretty', 'visitDurationPretty', 'referrerName', 'referrerKeyword']
        # print(df.columns)

        # Find columns that start with "serverTimePretty (actionDetails"
        action_details_columns = [col for col in df.columns if 'serverTimePretty (actionDetails' in col]
        page_title_columns = [col for col in df.columns if 'pageTitle (actionDetails' in col]

        # Initialize an empty DataFrame for filtered data
        df_filtered = pd.DataFrame()

        # Add the specific columns to df_filtered
        df_filtered = df[specific_columns]

        # Loop through each 'pageTitle (actionDetails X)' and 'serverTimePretty (actionDetails X)' column pair
        for page_title_col, server_time_col in zip(page_title_columns, action_details_columns):
            # Filter rows where 'pageTitle (actionDetails X)' is 'ThankYou - GoFresh'
            condition = df[page_title_col] == 'ThankYou - GoFresh'

            # Create a new column in df_filtered that contains values from server_time_col where condition is True,
            # NaN otherwise
            df_filtered[server_time_col] = df.loc[condition, server_time_col]

        # Remove rows where all serverTimePretty columns are NaN
        df_filtered.dropna(subset=action_details_columns, how='all', inplace=True)

        # Remove columns where all values are NaN
        df_filtered.dropna(axis=1, how='all', inplace=True)

        # Save the filtered DataFrame to a new CSV file
        df_filtered.to_csv('filtered_data.csv', index=False)

    def process_qualtrics(self):
        qualtrics_df = pd.read_csv(self.qualtrics_file, delimiter=',', dtype={'IP Address (first two segments)': str})
        
        # Convert the 'Recorded Date' and 'Start Date' columns in Qualtrics to datetime
        qualtrics_df['Recorded Date'] = pd.to_datetime(qualtrics_df['Recorded Date'], format='%m/%d/%Y %H:%M')
        qualtrics_df['Start Date'] = pd.to_datetime(qualtrics_df['Start Date'], format='%m/%d/%Y %H:%M')

        # Truncate seconds and keep up to minutes
        qualtrics_df['Recorded Date'] = qualtrics_df['Recorded Date'].dt.floor('T')

        # Add empty columns and new columns for matching
        for _ in range(4):  # Adding 4 empty columns
            qualtrics_df['                 '] = np.nan

        # Add 'IP and TimeStamp Matched/Not Matched' column and 'Repeated User' column
        qualtrics_df['TimeStamp Matched/Not Matched'] = 'Not Matched'
        qualtrics_df['IP Matched/Not Matched'] = 'Not Matched'
        qualtrics_df['Repeated User'] = 'No'

        for _ in range(4):  # Adding 4 more empty columns
            qualtrics_df['                '] = np.nan

        return qualtrics_df



    def join_data(self, qualtrics_df):
        # Load the filtered_data.csv into a DataFrame
        filtered_df = pd.read_csv('./filtered_data.csv')

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
                    matching_rows = filtered_df[filtered_df[col] == row['Recorded Date']]
                    if not matching_rows.empty:
                        qualtrics_df.at[index, 'TimeStamp Matched/Not Matched'] = 'Matched'
                        for new_col in new_columns:
                            qualtrics_df.at[index, new_col] = matching_rows.iloc[0][new_col]
        
        # Check if each truncated IP in qualtrics_df is a prefix in filtered_data.csv
        for index, row in qualtrics_df.iterrows():
            ip_prefix = str(row['IP Address (first two segments)'])  # Convert to string
            if any(str(ip).startswith(ip_prefix) for ip in filtered_ips):  # Convert IP to string before using startswith
                qualtrics_df.at[index, 'IP Matched/Not Matched'] = 'Matched'

        # Find repeated users based on IP
        count_by_ip = filtered_df.filter(like='serverTimePretty (actionDetails').apply(lambda x: x.count(), axis=1)
        repeated_users = filtered_df.loc[count_by_ip >= 2, 'visitIp'].unique()

        qualtrics_df['IP Address (first two segments)'] = qualtrics_df['IP Address (first two segments)'].astype(str)

        for ip in repeated_users:
            pattern = r'^' + re.escape(ip.split('.')[0]) + r'\.' + re.escape(ip.split('.')[1])
            qualtrics_df.loc[qualtrics_df['IP Address (first two segments)'].str.match(pattern, case=False), 'Repeated User'] = 'Yes'

        # To determine combined matching for timestamp and IP
        qualtrics_df['IP & TimeStamp Combined'] = np.where(
            (qualtrics_df['TimeStamp Matched/Not Matched'] == 'Matched') &
            (qualtrics_df['IP Matched/Not Matched'] == 'Matched'),
            'Matched', 'Not Matched'
        )

        # Sort rows by 'StartDate'
        qualtrics_df.sort_values(by='Start Date', inplace=True)

        # Replace NaN values with an empty string
        qualtrics_df = qualtrics_df.fillna('')

        return qualtrics_df
    
    def generate_output(self, qualtrics_df):

        # Reformat the 'Start Date' and 'Recorded Date' columns
        qualtrics_df['Start Date'] = qualtrics_df['Start Date'].dt.strftime('%m/%d/%Y %H:%M')
        qualtrics_df['Recorded Date'] = qualtrics_df['Recorded Date'].dt.strftime('%m/%d/%Y %H:%M')

        # Rename columns
        column_mapping = {
            'IP Address (first two segments)': 'IP Address',
            'TimeStamp Matched/Not Matched': 'Timestamp Match',
            'IP Matched/Not Matched': 'IP Match',
            'idVisit': 'Visit ID',
            'visitIp': 'Visit IP',
            'serverTimePretty': 'Server Time',
            'visitDurationPretty': 'Visit Duration',
            'referrerName': 'Referrer Name',
            'referrerKeyword': 'Referrer Keyword',
            'IP & TimeStamp Combined': 'IP & Timestamp Match'
        }
        qualtrics_df.rename(columns=column_mapping, inplace=True)

        # Open a new Excel file and add worksheets.
        workbook = xlsxwriter.Workbook(self.output_file, {'nan_inf_to_errors': True})

        
         # Create a new worksheet for column descriptions
        description_worksheet = workbook.add_worksheet('Column Descriptions')

        # Define column titles and their descriptions
        column_descriptions = {
            'Start Date': 'The date and time when the survey was started',
            'End Date': 'The date and time when the survey ended',
            'IP Address': 'The IP address of the survey participant',
            'Recorded Date': 'The date and time when the survey data was recorded',
            'Referral Source': 'The source from which the survey participant was referred',
            'Timestamp Match': 'Indicates whether the survey timestamp matches with other data',
            'IP Match': 'Indicates whether the survey IP address matches with other data',
            'Repeated User': 'Indicates if the user is a repeated visitor based on IP',
            'Visit ID': 'A unique identifier for the visit',
            'Visit IP': 'The IP address associated with the visit',
            'Server Time': 'The time when a server action occurred during the visit',
            'Visit Duration': 'The duration of the visit',
            'Referrer Name': 'The name of the referrer (source) that led to the visit',
            'Referrer Keyword': 'The keyword used by the referrer to access the site',
            'IP & Timestamp Match': 'Indicates whether both IP and timestamp match with other data'
        }

        # Define cell formatting for the title and descriptions
        title_format = workbook.add_format({
            'bold': True,
            'font_size': 14,
            'align': 'center',
            'valign': 'vcenter',
            'bg_color': '#D7E4BC',
            'border': 1,
            'text_wrap': True
        })

        description_format = workbook.add_format({
            'font_size': 12,
            'align': 'left',
            'valign': 'vcenter',
            'border': 1,
            'text_wrap': True
        })

        # Set the title and description for the "Column Descriptions" sheet
        title = 'Column Descriptions'
        description = 'This sheet provides descriptions for each column in the "Data" sheet.'
        description_worksheet.merge_range('A1:B1', title, title_format)
        description_worksheet.merge_range('A2:B2', description, description_format)

        # Increase column width for the "Column Descriptions" sheet
        description_worksheet.set_column('A:A', 20)  # Set column A to a width of 20
        description_worksheet.set_column('B:B', 60)  # Set column B to a width of 60


        # Write column titles and descriptions to the description worksheet
        description_worksheet.write(4, 0, 'Column Name', title_format)
        description_worksheet.write(4, 1, 'Description', title_format)
        for idx, (col_title, col_description) in enumerate(column_descriptions.items(), start=5):
            description_worksheet.write(idx, 0, col_title, description_format)
            description_worksheet.write(idx, 1, col_description, description_format)

        # Worksheet for the data
        data_worksheet = workbook.add_worksheet('Data')
        header_format = workbook.add_format({
            'bold': True,
            'fg_color': '#D7E4BC',
            'border': 1,
            'text_wrap': True,
            'valign': 'vcenter',
            'align': 'center'
        })
        cell_format = workbook.add_format({'border': 1})

        # Write DataFrame data to the data worksheet
        for row_num, (index, row_data) in enumerate(qualtrics_df.iterrows()):
            for col_num, cell_data in enumerate(row_data):
                data_worksheet.write(row_num + 1, col_num, cell_data, cell_format)

        # Apply column headers and formatting
        for col_num, column_name in enumerate(qualtrics_df.columns):
            data_worksheet.write(0, col_num, column_name, header_format)

        # Adjust column width to fit data
        for i, column in enumerate(qualtrics_df):
            max_len = max(
                qualtrics_df[column].astype(str).apply(len).max(),
                len(str(column)
            ))
            data_worksheet.set_column(i, i, max_len)

        # Define user-friendly colors
        chart_colors = ['#3CB371', '#FFA500']

        # Pie chart data
        categories = {
            'Timestamp Match': qualtrics_df['Timestamp Match'].value_counts().reindex(['Matched', 'Not Matched']).fillna(0),
            'IP Match': qualtrics_df['IP Match'].value_counts().reindex(['Matched', 'Not Matched']).fillna(0),
            'IP & Timestamp Match': qualtrics_df['IP & Timestamp Match'].value_counts().reindex(['Matched', 'Not Matched']).fillna(0)
        }

        # Create a new worksheet for the pie chart
        pie_chart_worksheet = workbook.add_worksheet('Charts')

        # Start column positions for the pie chart
        start_col = 0

        for idx, (title, values) in enumerate(categories.items()):
            end_row = len(values) + 1

            data_start_col = start_col + 2 * idx
            data_end_col = data_start_col + 1

            # Add the title as the column heading in the pie chart aggregation table
            pie_chart_worksheet.write(0, data_start_col, title)

            # Insert data for chart on the pie chart worksheet
            for row_num, (label, value) in enumerate(values.items(), start=1):
                pie_chart_worksheet.write(row_num, data_start_col, label)
                pie_chart_worksheet.write(row_num, data_end_col, value)

            # Create a new pie chart object.
            pie_chart = workbook.add_chart({'type': 'pie'})

            # Configure the chart with a title defined by the category.
            pie_chart.add_series({
                'name': title,
                'categories': ['Charts', 1, data_start_col, end_row, data_start_col],
                'values': ['Charts', 1, data_end_col, end_row, data_end_col],
                'points': [{'fill': {'color': chart_colors[i]}} for i in range(len(values))],
                'data_labels': {'percentage': True},
            })

            # Insert the pie chart on the pie chart worksheet
            chart_position_row = 0
            chart_position_col = 8 + idx
            pie_chart_worksheet.insert_chart(chart_position_row, chart_position_col, pie_chart)

        # Close the workbook
        workbook.close()

if __name__ == "__main__":

    # input the following data
    date = "10-13-2023" # insert date for reporting
    qualtrics_filename = './qualtrics.csv' # qualtrics csv data goes here
    matomo_data_filename = './matomo.csv' # matomo csv data goes here
    output_filename = f'Qualtric_DataValidation_{date}.xlsx' # name of the excel workbook file

    # Creates an instance of the DataProcessor class, specifying input and output file names.
    data_processor = DataProcessor(qualtrics_filename, matomo_data_filename, output_filename)
    
    # Processes Matomo data and perform necessary operations.
    data_processor.process_matomo()
    
    # Processes Qualtrics data and store it in a DataFrame.
    qualtrics_df = data_processor.process_qualtrics()

    # Joins the processed Qualtrics data with other data sources if needed.
    qualtrics_df2 = data_processor.join_data(qualtrics_df)

    # Generates the final report in an Excel file based on the processed data.
    data_processor.generate_output(qualtrics_df2)

    # Prints a message indicating that the report is ready with the specified output filename.
    print("Your report is ready with filename: ", output_filename)


# To run this code, follow these steps:
# 1. Ensure you have the required data files for your reporting.
#    - Qualtrics filename: Ensure it is located in the same directory as the Python file.
#    - Matomo filename: Ensure it is located in the same directory as the Python file.
# 2. Open your terminal and navigate to the directory where this code is located using the 'cd' command.
#    Example: 'cd foldername'
# 3. Run the code by entering the following command in the terminal:
#    'python nameofpythonfile.py'
# 4. Press 'Enter' to execute the code.
# 5. You will see the output file appear in the directory you are currently in.
# Go to the folder on your pc/mac and see the file.
