# Daily Revenue Impact Calculator

This Python script calculates the revenue impact for a specific date by comparing it to the average of the revenue values from the three previous weekdays. The script reads data from an Excel file named 'HourlyRevenue.xlsb', processes the data, and provides the revenue impact value.

## Table of Contents

- [Prerequisites](#prerequisites)
- [Usage](#usage)
- [Script Explanation](#script-explanation)

## Prerequisites

Before running this script, make sure you have the following prerequisites in place:

- Python installed on your system (Python 3.x is recommended).
- The following Python libraries:
  - `datetime`
  - `timedelta`
  - `pandas`

## Usage 

1. Replace 'HourlyRevenue.xlsb' with your desired Excel file in the same directory as this script.

2. Open your terminal or command prompt.

3. Navigate to the directory containing the script.

4. Run the script by executing the following command:

   ```
   python script.py
   ```

   Replace `script.py` with the name of the script file if it's different.

5. The script will prompt you to enter a target date. Uncomment the line that defines the target date and set it to your desired date (e.g., `date = datetime(2023, 8, 4)`).

6. Or you can import the script as a module into another file, for easier use. 

7. The script will calculate the revenue impact and display the result.

## Script Explanation

The script performs the following steps:

1. Imports the necessary libraries, including `datetime`, `timedelta`, and `pandas`.

2. Reads data from an Excel file named 'HourlyRevenue.xlsb' into a Pandas DataFrame.

3. Defines a function named `Revenue_impact` that takes a target date as input and calculates the revenue impact.

4. Calculates the list of previous weekdays (up to four weeks ago including the impact day) based on the target date.

5. Defines a helper function, `date_converter`, to convert dates to a specific format with week numbers.

6. Converts the previous weekdays to the desired format with week numbers.

7. Extracts specific data values from the DataFrame for the target date and the three previous weekdays.

8. Calculates the average of the extracted values.

9. Calculates and prints the revenue impact by subtracting the average value from the target date's value.

The script is a useful tool for analyzing the impact of daily revenue fluctuations and can be customized to work with different datasets by adjusting the Excel file's format and data extraction methods.
