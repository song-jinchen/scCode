import pandas as pd
import csv

# Load spreadsheet
xl = pd.ExcelFile('timetable0801-0802.xls')

# Load a sheet into a DataFrame by its name
df = xl.parse('Sheet1')

# Write the DataFrame to output.csv
df.to_csv('output.csv', index=False, encoding='utf-8')

# Create the timetable.csv file
with open('timetable.csv', mode='w', newline='', encoding='utf-8') as file:
    writer = csv.writer(file)

    # Write the header for dates 1 and 2
    header = ["姓名"]
    for i in range(1, 3):  # Change the range to include dates 1 to 2 (i.e., 1, 2)
        header.extend([str(i) + "上班时间", str(i) + "下班时间", ""])
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

            # Write the data
            data_row = [name]
            for time_pair in range(0, len(times_formatted), 2):
                data_row.extend([times_formatted[time_pair], times_formatted[time_pair + 1], ""]) # Add an empty column for hours
            writer.writerow(data_row)

            index += 3  # Move to the next chunk of 3 rows
