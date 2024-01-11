import pandas as pd
from datetime import datetime, timedelta  # Import timedelta from the datetime module


file = 'excel.xlsx'

data = pd.read_excel(file)


# code for part a 

df = pd.DataFrame(data)
df['Time'] = pd.to_datetime(df['Time'])
df['Time Out'] = pd.to_datetime(df['Time Out'])

# Create a new DataFrame with a list of dates for each employee
df_dates = df.groupby('Employee Name')['Time'].apply(list).reset_index(name='Dates')

# Check for consecutive days attendance
df_dates['Consecutive Days'] = df_dates['Dates'].apply(lambda dates: any((date - pd.DateOffset(days=i)) in dates for date in dates for i in range(6)))

# Filter the DataFrame to get employees who attended for 7 consecutive days
result = df_dates[df_dates['Consecutive Days']][['Employee Name']]

# Merge with the original DataFrame to get 'Position ID'
result = pd.merge(result, df[['Employee Name', 'Position ID']].drop_duplicates(), on='Employee Name')

# Display the result
print ("part A")
print(result)


# code for part b

df = pd.DataFrame(data)

# Convert 'Time Out' column to datetime format
df['Time Out'] = pd.to_datetime(df['Time Out'], format='%m/%d/%Y %I:%M %p')

# Calculate time difference between shifts
df['Time Difference'] = df['Time Out'].diff()

# Filter employees with less than 10 hours of time between shifts but greater than 1 hour
result = df[(df['Time Difference'] < timedelta(hours=10)) & (df['Time Difference'] > timedelta(hours=1))]

# Drop duplicate entries based on employee information
result = result.drop_duplicates(subset=['Position ID', 'Employee Name'])

# Display Employee ID and Employee Name
print ("Part B")
print(result[['Position ID', 'Employee Name']])


# code for part c 

df = pd.DataFrame(data)

# Convert 'Timecard Hours (as Time)' to total hours in float format with error handling
def convert_to_hours(x):
    try:
        return sum(int(i) * 60**index for index, i in enumerate(reversed(x.split(':'))))/60
    except (AttributeError, ValueError):
        return None

df['Timecard Hours'] = df['Timecard Hours (as Time)'].apply(convert_to_hours)

# Filter employees who have worked for more than 14 hours in a single shift
result = df[df['Timecard Hours'] > 14]

# Display Employee ID and Employee Name
print("part C")
print(result[['Position ID', 'Employee Name']])