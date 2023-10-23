# Matomo & Qualtrics Data Matching

## Description

This project is aimed at processing and matching data from Matomo and Qualtrics. It contains three Python scripts:

1. `app.py`: Entry point of the application, which calls the other two processes.
2. `matomo_processing.py`: For processing Matomo data, filtering specific columns, and saving the results to a new CSV file.
3. `qualtrics_matching.py`: For matching Qualtrics data with filtered Matomo data, and generating an Excel file with additional columns and visualizations.

## Prerequisites

- Python 3.x
- pandas library
- numpy library
- xlsxwriter library

## Installation

1. Clone the repository:

    ```bash
    git clone https://github.com/nikhilchoppa/QualtricsDataValidation.git
    ```

2. Install the required packages:

    ```bash
    pip install pandas numpy xlsxwriter
    ```

## Usage

### Running the Application

1. Run `app.py`:

    ```bash
    python app.py
    ```

    This will trigger both `matomo_processing.py` and `qualtrics_matching.py`.

### Input Files

1. Place your Matomo CSV file in the project directory. The default name should be 'Matomo.csv'.
2. Place your Qualtrics CSV file in the project directory. The default name should be 'Qualtrics.csv'.

### Output Files

1. `filtered_matomo_data.csv`: This will be generated by `matomo_processing.py`.
2. `Qualtrics_Validation_All_10-03To10-16.xlsx`: This will be generated by `qualtrics_matching.py`.

## Scripts

### `app.py`

This script serves as the entry point of the application. It calls `run_matomo_process` and `run_qualtrics_matching` functions.

**Key Functions:**

- `run_matomo_process(input_file)`: Executes `matomo_processing.py` with the given input file.
- `run_qualtrics_matching(input_csv, output_xlsx)`: Executes `qualtrics_matching.py` with the given input CSV and output Excel file.

### `matomo_processing.py`

This script processes Matomo data.

**Key Functions:**

- `process_matomo(input_file)`: Reads the Matomo CSV, filters out specific columns and values, and saves it to a new CSV file.

### `qualtrics_matching.py`

This script performs Qualtrics data matching.

**Key Functions:**

- `process_qualtrics(input_csv, output_xlsx)`: Reads the Qualtrics CSV, matches with filtered Matomo data, and saves it to an Excel file.


