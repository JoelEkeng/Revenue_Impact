# Import necessary libraries
from datetime import datetime, timedelta
import pandas as pd

# Read the Excel file into a DataFrame
df = pd.read_excel('HourlyRevenue.xlsb')

# Define the target date
#date = datetime(2023, 8, 4)

#The main function
def Revenue_impact(date):
    # Calculate the previous week Days including the target date
    previous_weekdays = []
    for _ in range(4):
        target_date = date  # Reset target_date to the original date
        while 7<= target_date.weekday() != 0:  
            target_date -= timedelta(days=1)
        previous_weekdays.append(target_date.strftime('%Y-%m-%d'))
        date -= timedelta(days=7)  # Move to the previous week

    # Print the list of previous weekdays
    print("Previous Weekdays:", previous_weekdays)

    # Define a function to determine the week number based on the day of the month
    def date_converter(date):
        day = int(date.split("-")[2])  # Extract day from the date string


        if 22 <= day >= 31:
            Week = 4
        else: 
            Week = (day - 1) // 7 + 1  # Calculate week number

        inner_week = (day - 1) % 7  # Calculate inner week number
        if inner_week == 0:
            # Convert the date to the desired format
            long_string_date = datetime.strptime(date, '%Y-%m-%d').strftime("%b-%y Wk {}".format(Week))
        
            return long_string_date
        else:
            # Convert the date to the desired format
            long_string_date = datetime.strptime(date, '%Y-%m-%d').strftime("%b-%y Wk {}.{}".format(Week, inner_week))
            
            return long_string_date

    # Convert the previous weekdays to long string format with week numbers
    long_string_dates = [date_converter(prev_weekday) for prev_weekday in previous_weekdays]

    #Incidence captured for analysis 
    Incident = long_string_dates[0]
    Previous_day_1 = long_string_dates[1]
    Previous_day_2 = long_string_dates[2]
    Previous_day_3 = long_string_dates[3]
    print("Long dates: {}, {}, {}, {}".format(Incident,Previous_day_1,Previous_day_2,Previous_day_3))

    # Extract specific data values from the DataFrame
    exactday = df.at[31, Incident]
    day1 = df.at[31, Previous_day_1]
    day2 = df.at[31, Previous_day_2]
    day3 = df.at[31, Previous_day_3]

    # Calculate the average of the extracted values
    average_total = (day1 + day2 + day3) / 3

    # Calculate and print the revenue impact
    revenue_impact_value = exactday - average_total
    print("Revenue Impact:", revenue_impact_value)
