import pandas as pd
import csv
from datetime import datetime, timedelta

def mark_time(time_str, is_start_time=True):
    if not time_str or time_str == ":00":
        return ""

    time_obj = datetime.strptime(time_str, "%H:%M:%S")
    if is_start_time:
        if time_obj.time() < datetime.strptime("06:35:00", "%H:%M:%S").time():
            return "06:30:00"
        elif time_obj.time() < datetime.strptime("07:00:00", "%H:%M:%S").time():
            return "07:00:00"
        else:
            # Round up to the nearest half hour
            minute = time_obj.minute
            if minute < 30:
                rounded_minute = 30
            else:
                rounded_minute = 0
            return f"{time_obj.hour:02d}:{rounded_minute:02d}:00"
    else:
        if time_obj.time() < datetime.strptime("15:25:00", "%H:%M:%S").time():
            return f"{time_obj.hour:02d}:30:00"
        elif time_obj.time() > datetime.strptime("15:54:00", "%H:%M:%S").time():
            return f"{time_obj.hour:02d}:{time_obj.minute // 30 * 30:02d}:00"
        else:
            # Round down to the nearest half hour
            minute = time_obj.minute
            rounded_minute = (minute // 30) * 30
            return f"{time_obj.hour:02d}:{rounded_minute:02d}:00"

# Load spreadsheet
xl = pd.ExcelFile('考勤记录表_8月.xlsx')

# Load a sheet into a DataFrame by its name
df = xl.parse('Sheet1')

# Write the DataFrame to output.csv
df.to_csv('output.csv', index=False, encoding='utf-8')

# Create the timetable.csv file
with open('timetable.csv', mode='w', newline='', encoding='utf-8') as file:
    writer = csv.writer(file)

    # Write the header for dates 1 and 2, and the new "hours" column
    header = ["姓名"]
    for i in range(1, 3):  # Change the range to include dates 1 to 2 (i.e., 1, 2)
        header.extend([str(i) + "上班时间", str(i) + "下班时间", "hours"])
    writer.writerow(header)

    # Open the CSV file that we just created
    with open('output.csv', newline='', encoding='utf-8') as csvfile:
        reader = csv.reader(csvfile)
        rows = list(reader)

        # Initialize variables to store employee information
        name = ''

        # Iterate through the rows
        index = 0
        while index < len(rows) - 2:  # Make sure there are enough rows remaining
            # Check if the row contains '工号'
            if '工号' not in rows[index][0]:
                index += 1
                continue

            # Extract information from the first row (姓名)
            info_row = rows[index]
            for item in info_row:
                if "姓名" in item:
                    name = item.split("：")[1]

            # Extract times
            times_row = rows[index + 2][:7]
            times_formatted = []
            for time in times_row:
                if "\n" in str(time):
                    start_end = str(time).split("\n")
                    times_formatted.extend([start_end[0] + ":00", start_end[1] + ":00"])
                else:
                    times_formatted.extend(["", ""])  # If there's no time, add two empty strings

            # Calculate working hours for each day
            working_hours = []
            for i in range(0, len(times_formatted), 2):
                start_time = mark_time(times_formatted[i], is_start_time=True)
                end_time = mark_time(times_formatted[i + 1], is_start_time=False)

                if start_time and end_time:
                    start_time_obj = datetime.strptime(start_time, "%H:%M:%S")
                    end_time_obj = datetime.strptime(end_time, "%H:%M:%S")
                    working_time = end_time_obj - start_time_obj
                    working_hours.append(str(working_time))
                else:
                    working_hours.append("")  # If either start or end time is missing, add an empty string

            # Write the data
            data_row = [name]
            for i in range(0, len(times_formatted), 2):
                data_row.extend([times_formatted[i], times_formatted[i + 1], working_hours[i // 2]])
            writer.writerow(data_row)

            index += 3  # Move to the next chunk of 3 rows