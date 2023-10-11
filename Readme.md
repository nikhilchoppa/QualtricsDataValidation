# Data Validation: Qualtrics and Matomo Integration

This project comprises scripts designed to facilitate the validation of Qualtrics datasets by comparing them against Matomo's "thankyou-social" submissions. The scripts read, process, analyze the datasets, and generate an Excel file with visual representation in the form of pie charts.

## Files in the Project:

### 1. `matomo_processing.py`: 
- **Description**: This script focuses on refining the raw Matomo data. It reads the dataset and filters out irrelevant entries, particularly focusing on the "thankyou-social" submissions. The result is a streamlined CSV that makes the matching process with the Qualtrics data more efficient.
  
- **Main Functions**: 
  - Data Filtering: Removes rows not related to "thankyou-social".
  - Extraction of Relevant Columns: Retains only the columns that are essential for the matching process.

### 2. `qualtrics_matching.py`: 
- **Description**: The core of the project, this script takes the Qualtrics dataset and the pre-processed Matomo data, performs an intricate matching operation, and then visually represents the results in an Excel file. The visual representation aids in quickly understanding the proportion of data that matches between the two sources.
  
- **Main Functions**:
  - Data Matching: Compares Qualtrics and Matomo data based on specific parameters.
  - Excel Output Generation: Constructs an Excel sheet complete with pie charts visualizing the matched and unmatched data.
  - Data Visualization: Uses xlsxwriter to craft detailed pie charts that visually represent the results of the matching operation.

## Prerequisites:
- Ensure you have Python 3.10 installed.
- Required Libraries:
  - pandas
  - numpy
  - xlsxwriter

## How to Use:

1. **Setting up the Environment**
   - Install the necessary Python packages using pip:
     ```
     pip install pandas numpy xlsxwriter
     ```

2. **Processing the Matomo Dataset**
   - Place the `Matomo.csv` file in your working directory.
   - Run the `matomo_processing.py` script:
     ```
     python matomo_processing.py
     ```
   - This will generate a `filtered_data.csv` that contains only relevant rows from the Matomo dataset.

3. **Matching and Analyzing the Qualtrics Data**
   - Ensure `Qualtrics1.csv` is in your working directory.
   - Run the `qualtrics_matching.py` script:
     ```
     python qualtrics_matching.py
     ```
   - This will output an Excel file named `Qualtrics_Validation_Facebook_09-11To09-22.xlsx` with the comparison results, including pie charts visualizing the matching data.

## Project Highlights:

- **Data Filtering and Validation**: The scripts efficiently process raw data, filtering out unnecessary rows and columns, ensuring that only relevant data is utilized for matching and analysis.
  
- **Visualization**: The results are visually represented in the form of pie charts in the Excel output, making it easy to discern the proportion of matched vs unmatched data.

- **Efficiency**: Through Python's powerful libraries, the scripts handle large datasets smoothly, ensuring optimal performance.

## Future Enhancements:

- Integrate more visualization techniques, such as bar charts or histograms.
- Add functionality to automate the fetching of datasets from respective sources.
- Implement logging to keep track of script operations and any potential issues.

## Feedback and Contributions:

We appreciate feedback and contributions. If you find any issues or have suggestions for improvements, please open an issue or submit a pull request.



