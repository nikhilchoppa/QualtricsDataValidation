# Data Processing and Matching Scripts

This repository contains two primary Python scripts dedicated to processing and matching data sets:

1. **Matomo Data Processor (`matomo_processing.py`)**: A tool that filters a dataset (`Matomo.csv`) to capture specific visitor interactions, saving the output to `filtered_data.csv`.
2. **Qualtrics Data Matcher (`qualtrics_matching.py`)**: This script matches survey responses from `Qualtrics.csv` against the processed data in `filtered_data.csv`, identifying and annotating matches, resulting in a combined dataset saved as `Qualtrics_matching.csv`.

## Detailed Descriptions

### 1. Matomo Data Processor (`matomo_processing.py`)
- **Purpose**: Designed to sieve through a dataset, focusing on specific interactions related to the 'ThankYou - GoFresh' page title.
  
- **Key Features**:
  - **Data Loading**: Imports data from 'Matomo.csv', keeping in mind the tab delimiters and utf-16 encoding.
  - **Column Filtering**: Specifically seeks out columns of interest and those related to action details.
  - **Data Filtration**: Filters rows based on the values in the 'pageTitle (actionDetails X)' columns.
  - **Cleanup**: The script omits NaN-heavy rows and columns to refine the dataset.
  - **Output**: The polished data is stored in `filtered_data.csv` for subsequent use.

### 2. Qualtrics Data Matcher (`qualtrics_matching.py`)
- **Purpose**: Efficiently processes and matches survey data from Qualtrics with a previously processed dataset, merging the data and flagging matches and repeated users.
  
- **Key Features**:
  - **Date Formatting**: It loads 'Qualtrics.csv' and tailors the format of date columns for further processing.
  - **Data Matching**: By utilizing the 'RecordedDate' field, it matches survey data with entries in `filtered_data.csv`.
  - **Repeated User Identification**: It flags users who have appeared multiple times based on IP address analysis.
  - **Output**: The consolidated and matched data is captured in `Qualtrics_matching.csv`.

## Prerequisites
1. Ensure that Python is installed on your machine.
2. Install the essential libraries:
   ```bash
   pip install pandas numpy
   ```

## Usage
1. Place the required CSV files (`Matomo.csv` and `Qualtrics.csv`) in the working directory before executing the scripts.
2. Initiate the Matomo data processing:
   ```bash
   python matomo_processing.py
   ```
   This creates `filtered_data.csv`.
3. Next, run the Qualtrics data matcher:
   ```bash
   python qualtrics_matching.py
   ```
   This synthesizes `Qualtrics_matching.csv` using the previously generated `filtered_data.csv` and original `Qualtrics.csv`.

## Contributing
Your contributions can further refine these scripts. Please initiate an issue for discussions or directly submit a pull request.

