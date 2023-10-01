import openpyxl
from datetime import datetime, timedelta

# Function to calculate the difference in hours between two time strings
def calculate_hour_difference(time_str1, time_str2):
    time_format = "%H:%M:%S"
    time1 = datetime.strptime(time_str1, time_format)
    time2 = datetime.strptime(time_str2, time_format)
    time_difference = abs((time2 - time1).total_seconds() / 3600)  # Convert seconds to hours
    return time_difference

# Load the Excel file
file_path = r"C:\Users\Vivek\Desktop\Project\Assignment_Timecard.xlsx"
workbook = openpyxl.load_workbook(file_path)
sheet = workbook.active

# Define dictionaries to store employee data
employee_data = {}

# Loop through rows in the Excel sheet
for row in sheet.iter_rows(min_row=2, values_only=True):
    if len(row) >= 4:
        name, position, date_str, start_time, end_time, *extra_columns = row

        if name not in employee_data:
            employee_data[name] = []

        # Check if the date string and time strings are not empty
        if date_str and start_time and end_time and isinstance(date_str, str) and isinstance(start_time, str) and isinstance(end_time, str):
            try:
                # Assuming date_str is in a format like 'MM/DD/YYYY'
                date_obj = datetime.strptime(date_str, "%m/%d/%Y")
                start_time_obj = datetime.strptime(start_time, "%H:%M:%S")
                end_time_obj = datetime.strptime(end_time, "%H:%M:%S")
                employee_data[name].append((date_obj, start_time_obj, end_time_obj, position))
            except ValueError:
                print(f"Skipping invalid date or time: {date_str}, {start_time}, {end_time}")

# Analyze employee data
for name, shifts in employee_data.items():
    shifts.sort(key=lambda x: x[0])  # Sort shifts by date
    consecutive_days = 1
    total_hours = 0

    for i in range(1, len(shifts)):
        prev_shift, current_shift = shifts[i - 1], shifts[i]
        days_difference = (current_shift[0] - prev_shift[0]).days
        time_difference = calculate_hour_difference(prev_shift[2].strftime("%H:%M:%S"), current_shift[1].strftime("%H:%M:%S"))

        if days_difference == 1:
            consecutive_days += 1
            total_hours += time_difference
        else:
            consecutive_days = 1
            total_hours = 0

        if consecutive_days >= 7:
            print(f"{name} ({current_shift[3]}) has worked for 7 consecutive days.")

        if 1 < time_difference < 10:
            print(f"{name} ({current_shift[3]}) has less than 10 hours between shifts but greater than 1 hour.")

        if time_difference > 14:
            print(f"{name} ({current_shift[3]}) has worked for more than 14 hours in a single shift.")

# Close the Excel file
workbook.close()
