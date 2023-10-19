import pandas as pd
import sys

def process_matomo(input_file):

    # Read the CSV file into a DataFrame with the correct delimiter and encoding
    df = pd.read_csv(input_file, delimiter='\t', encoding='utf-16')

    # Create a list of specific columns you want to keep
    specific_columns = ['idVisit', 'visitIp', 'serverTimePretty', 'visitDurationPretty', 'referrerName', 'referrerKeyword']

    # Find columns that start with "serverTimePretty (actionDetails"
    action_details_columns = [col for col in df.columns if 'serverTimePretty (actionDetails' in col]
    page_title_columns = [col for col in df.columns if 'pageTitle (actionDetails' in col]

    # Initialize an empty DataFrame for filtered data
    df_filtered = pd.DataFrame()

    # Add the specific columns to df_filtered
    df_filtered[specific_columns] = df[specific_columns]

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

    # Terminate the program
    sys.exit(0)


if __name__ == "__main__":
    if len(sys.argv) > 1:
        input_file = sys.argv[1]
        process_matomo(input_file)
    else:
        process_matomo()