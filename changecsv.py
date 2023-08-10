import pandas as pd

# Load spreadsheet
xl = pd.ExcelFile('timetable0719-0725.xls')

# Load a sheet into a DataFrame by its name
df = xl.parse('Sheet1')

# Prepare header
header = ["姓名", "工号", "部门"]
for i in range(19, 26):
    header.extend([str(i), "", ""])

# Create a DataFrame to store the timetable
timetable_df = pd.DataFrame(columns=header)

# Helper function to extract worker info
def process_worker_info(row, times_row):
    employee_number = row[0].split("：")[1]
    name = row[2].split("：")[1]
    department = row[4].split("：")[1]

    times_formatted = []
    for time in times_row:
        if "\n" in str(time):
            start_end = str(time).split("\n")
            times_formatted.extend([start_end[0] + ":00", start_end[1] + ":00", ""])
        else:
            times_formatted.extend(["", "", ""])

    return [name, employee_number, department] + times_formatted

# Iterate through the rows and process worker information
i = 0
while i < len(df):
    row = df.loc[i]
    # If we find "工号", it's a worker info row
    if '工号' in str(row[0]):
        # Handle first worker in the unit
        first_worker_info = process_worker_info(row, df.loc[i + 1][:7])
        timetable_df.loc[len(timetable_df)] = first_worker_info

        # Handle second worker in the unit
        second_worker_info = process_worker_info(df.loc[i + 3], df.loc[i + 5][:7])
        timetable_df.loc[len(timetable_df)] = second_worker_info

        i += 6  # Skip to next unit
    else:
        i += 1

# Write the DataFrame to timetable.csv
timetable_df.to_csv('timetable.csv', index=False, encoding='utf-8')
