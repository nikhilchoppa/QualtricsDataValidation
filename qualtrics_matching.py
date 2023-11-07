import pandas as pd
import numpy as np
import sys
import xlsxwriter


def process_qualtrics(input_csv, output_xlsx):
    # Load the Qualtrics CSV into a DataFrame and specify 'IPAddress' as a string
    qualtrics_df = pd.read_csv(input_csv, delimiter=',', dtype={'IPAddress': str})

    # Convert the 'RecordedDate' and 'StartDate' columns in Qualtrics to datetime
    qualtrics_df['RecordedDate'] = pd.to_datetime(qualtrics_df['RecordedDate'], format='%m/%d/%Y %H:%M',
                                                  errors='coerce')
    qualtrics_df['StartDate'] = pd.to_datetime(qualtrics_df['StartDate'], format='%m/%d/%Y %H:%M', errors='coerce')

    # Add 'IP and TimeStamp Matched/Not Matched' column and 'Repeated User' column
    qualtrics_df['TimeStamp Matched/Not Matched'] = 'Not Matched'
    qualtrics_df['IP Matched/Not Matched'] = 'Not Matched'
    qualtrics_df['Repeated User'] = 'No'

    # Load the filtered_matomo_data.csv into a DataFrame
    filtered_df = pd.read_csv('filtered_matomo_data.csv')

    # Get unique IPs from filtered_matomo_data.csv
    filtered_ips = filtered_df['visitIp'].unique()

    # Initialize new columns with NaN values
    new_columns = ['idVisit', 'visitIp', 'visitDurationPretty', 'referrerName', 'referrerKeyword']
    for col in new_columns:
        qualtrics_df[col] = np.nan

    # Loop through each 'serverTimePretty (actionDetails X)' column in filtered_matomo_data.csv
    for col in filtered_df.columns:
        if 'serverTimePretty (actionDetails' in col:
            # Convert the column to datetime
            filtered_df[col] = pd.to_datetime(filtered_df[col], format='%b %d, %Y %H:%M:%S')
            filtered_df[col] = filtered_df[col].dt.floor('T')

            # Update the 'Matched' column and new columns in qualtrics_df for matching rows
            for index, row in qualtrics_df.iterrows():
                time_lower_bound = pd.Timestamp(row['RecordedDate']) - pd.Timedelta(seconds=30)
                time_upper_bound = pd.Timestamp(row['RecordedDate']) + pd.Timedelta(seconds=30)

                matching_rows = filtered_df[
                    (filtered_df[col] >= time_lower_bound) & (filtered_df[col] <= time_upper_bound)]
                if not matching_rows.empty:
                    qualtrics_df.at[index, 'TimeStamp Matched/Not Matched'] = 'Matched'
                    qualtrics_df.at[index, 'ThankYou-ActionTime'] = matching_rows.iloc[0][col].strftime(
                        '%m/%d/%y %H:%M')

                    for new_col in new_columns:
                        qualtrics_df.at[index, new_col] = matching_rows.iloc[0][new_col]

    # Truncate IPs for comparison
    truncated_visit_ips = qualtrics_df['visitIp'].apply(lambda x: '.'.join(str(x).split('.')[:2]))
    truncated_ip_addresses = qualtrics_df['IPAddress'].apply(lambda x: '.'.join(str(x).split('.')[:2]))

    # Create 'IP Matched/Not Matched' column based on truncated IPs
    qualtrics_df['IP Matched/Not Matched'] = np.where(
        truncated_visit_ips == truncated_ip_addresses, 'Matched', 'Not Matched'
    )

    # Find repeated users based on 'idVisit'
    repeated_qualtrics_idvisits = qualtrics_df['idVisit'].value_counts().where(lambda x: x > 1).dropna().index.tolist()
    for idvisit in repeated_qualtrics_idvisits:
        qualtrics_df.loc[qualtrics_df['idVisit'] == idvisit, 'Repeated User'] = 'Yes'

    # To determine combined matching for timestamp and IP
    qualtrics_df['IP & TimeStamp Combined'] = np.where(
        (qualtrics_df['TimeStamp Matched/Not Matched'] == 'Matched') &
        (qualtrics_df['IP Matched/Not Matched'] == 'Matched'),
        'Both Matched', 'Not Both Matched'
    )

    # To determine combined matching for timestamp or IP
    qualtrics_df['IP or TimeStamp Either'] = np.where(
        (qualtrics_df['TimeStamp Matched/Not Matched'] == 'Matched') |
        (qualtrics_df['IP Matched/Not Matched'] == 'Matched'),
        'Either Matched', 'Neither Matched'
    )

    # Sort rows by 'StartDate'
    qualtrics_df.sort_values(by='StartDate', inplace=True)

    # Replace NaN values with an empty string
    qualtrics_df = qualtrics_df.fillna('')

    # Open a new Excel file and add a worksheet.
    workbook = xlsxwriter.Workbook(output_xlsx, {'nan_inf_to_errors': True})
    worksheet = workbook.add_worksheet('Qualtrics Data Validation')

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
            if pd.notna(cell_data):
                worksheet.write(row_num + 1, col_num, cell_data, cell_format)
            else:
                worksheet.write(row_num + 1, col_num, "N/A",
                                cell_format)  # Replace "N/A" with any default value you prefer

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

    # Add conditional formatting to highlight 'Yes' in the 'Repeated User' column with yellow
    repeated_user_col = qualtrics_df.columns.get_loc('Repeated User')  # Get the index of the 'Repeated User' column
    worksheet.conditional_format(1, repeated_user_col, len(qualtrics_df), repeated_user_col, {
        'type': 'text',
        'criteria': 'containing',
        'value': 'Yes',
        'format': workbook.add_format({'bg_color': '#FFFF00', 'font_color': '#000000'})  # Yellow background, black text
    })

    # Define pie chart colors
    chart_colors = ['#3CB371', '#FFA500']

    # Pie chart data
    categories = {
        'TimeStamp Matched/Not Matched': qualtrics_df['TimeStamp Matched/Not Matched'].value_counts().reindex(
            ['Matched', 'Not Matched']).fillna(0),
        'IP Matched/Not Matched': qualtrics_df['IP Matched/Not Matched'].value_counts().reindex(
            ['Matched', 'Not Matched']).fillna(0),
        'IP & TimeStamp Combined': qualtrics_df['IP & TimeStamp Combined'].value_counts().reindex(
            ['Both Matched', 'Not Both Matched']).fillna(0),
        'IP or TimeStamp Either': qualtrics_df['IP or TimeStamp Either'].value_counts().reindex(
            ['Either Matched', 'Neither Matched']).fillna(0)

    }

    # Create pie charts using xlsxwriter
    start_col = len(qualtrics_df.columns) + 2
    for idx, (title, values) in enumerate(categories.items()):
        end_row = len(values) + 1

        data_start_col = start_col + 4 * idx
        data_end_col = data_start_col + 1

        # Insert column titles for pie chart data
        worksheet.write(0, data_start_col, title)
        worksheet.write(1, data_start_col, "Status")
        worksheet.write(1, data_end_col, "Count")

        # Insert data for chart
        filtered_values = {k: v for k, v in values.items() if v > 0}  # Remove zero counts
        for row_num, (label, value) in enumerate(filtered_values.items(), start=2):
            worksheet.write(row_num, data_start_col, label)
            worksheet.write(row_num, data_end_col, value)

        # Create a new pie chart object.
        pie_chart = workbook.add_chart({'type': 'pie'})

        # Configure the chart, with a title defined by the category.
        pie_chart.add_series({
            'name': title,
            'categories': ['Qualtrics Data Validation', 2, data_start_col, end_row, data_start_col],
            'values': ['Qualtrics Data Validation', 2, data_end_col, end_row, data_end_col],
            'points': [{'fill': {'color': chart_colors[i]}} for i, _ in enumerate(filtered_values)],
            'data_labels': {'percentage': True},
        })
        pie_chart.set_title({'name': title})

        # Adjusted pie chart positioning
        chart_position_row = 0
        chart_position_col = data_end_col + 2
        worksheet.insert_chart(chart_position_row, chart_position_col, pie_chart)

    # Create a new worksheet for descriptions
    worksheet_description = workbook.add_worksheet("Descriptions")

    # Define Excel formats for headers and cells
    header_format = workbook.add_format({
        'bold': True,
        'font_size': 14,
        'fg_color': '#D7E4BC',
        'border': 1,
        'text_wrap': True,
        'valign': 'vcenter',
        'align': 'center'
    })

    cell_format = workbook.add_format({
        'font_size': 12,
        'border': 1,
        'text_wrap': True,
        'valign': 'vcenter'
    })

    # Populate the 'Descriptions' worksheet with code and column descriptions
    worksheet_description.write('A1', 'Column Descriptions:', header_format)

    # Initialize variables to hold maximum lengths for each column
    max_length_col_0 = 0
    max_length_col_1 = 0

    # Column descriptions
    columns_descriptions = {
        "StartDate": "The date and time when the Qualtrics survey was started by the User.",
        "EndDate": "The date and time when the Qualtrics survey ended by the User.",
        "IPAddress": "The Partial IP address of the Qualtrics survey participant.",
        "ReferralSource": "The sources from which the Qualtrics survey participant was referred.",
        "RecordedDate": "The date and time when the Qualtrics survey data was recorded.",
        "TimeStamp Matched/Not Matched": "Indicates if the timestamp was matched with Matomo data.",
        "IP Matched/Not Matched": "Indicates if the IP was matched with Matomo data.",
        "RepeatedUser": "Indicates if the user is a repeated visitor(multiple entries) based on VisitID.",
        "idVisit": "Unique identifier for each visit by Matomo.",
        "visitIp": "IP address of the visitor by Matomo.",
        "visitDurationPretty": "The time spent by User during visit by Matomo.",
        "referrerName": "The name of the referrer (source) that led to the visit by Matomo.",
        "referrerKeyword": "The keyword used by the referrer to access the site by Matomo.",
        "ThankYou-ActionTime": "The time when a Thank You action occurred during the visit by Matomo.",
        "IP & TimeStamp Combined": "Indicates if both IP and timestamp were matched.",
        "IP or TimeStamp Either": "Indicates if either IP or timestamp was matched."
    }

    row_num = 2  # Start row
    col_num = 0  # Start column

    for column, description in columns_descriptions.items():
        worksheet_description.write(row_num, col_num, column, cell_format)
        worksheet_description.write(row_num, col_num + 1, description, cell_format)

        # Update maximum lengths for columns if needed
        max_length_col_0 = max(max_length_col_0, len(column))
        max_length_col_1 = max(max_length_col_1, len(description))

        row_num += 1  # Move to the next row

    # Adjust column widths based on max lengths
    worksheet_description.set_column('A:A', max_length_col_0 + 2)  # +2 for a bit of padding
    worksheet_description.set_column('B:B', max_length_col_1 + 2)

    # Close the Pandas Excel writer and output the Excel file.
    workbook.close()
    sys.exit(0)


if __name__ == "__main__":
    if len(sys.argv) > 2:
        input_csv = sys.argv[1]
        output_xlsx = sys.argv[2]
        process_qualtrics(input_csv, output_xlsx)
    else:
        process_qualtrics()
