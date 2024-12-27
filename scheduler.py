import pandas as pd
import numpy as np
import random
from datetime import datetime, timedelta
import matplotlib.pyplot as plt
import matplotlib.dates as mdates
from collections import defaultdict



df = pd.read_excel('/content/Product Details_v2.xlsx')

# Convert columns to appropriate types
df['Promised Delivery Date'] = pd.to_datetime(df['Promised Delivery Date'])
df['Ready Time'] = pd.to_datetime(df['Ready Time'])
df['Start Time'] = pd.NaT  # Initialize as empty datetime
df['End Time'] = pd.NaT  # Initialize as empty datetime

# Sort the data by Promised Delivery Date, Product Name, and Component order
df = df.sort_values(by=['Promised Delivery Date',
                        'Product Name',
                        'Components']).reset_index(drop=True)

df[['Product Name','Promised Delivery Date',
    'Components','Machine Number','Operation',
    'Start Time','End Time']]


import pandas as pd
from datetime import datetime, timedelta
from collections import defaultdict

# Define working hours and working days
WORK_START = 9  # 9 AM
WORK_END = 17 + 1/60  # 5 PM
WEEKENDS = [5, 6]  # Saturday and Sunday

# Function to find the next valid working day
def next_working_day(current_date):
    while current_date.weekday() in WEEKENDS:  # Check if the day is a weekend
        current_date += timedelta(days=1)  # Move to the next day
    return current_date

# Function to adjust end time for working hours and working days
def adjust_to_working_hours_and_days(start_time, run_time_minutes):
    DAILY_WORK_MINUTES = (WORK_END - WORK_START) * 60  # Convert hours to minutes
    current_time = start_time
    remaining_minutes = run_time_minutes

    while remaining_minutes > 0:
        if current_time.hour < WORK_START or current_time.weekday() in WEEKENDS:
            current_time = next_working_day(current_time.replace(hour=WORK_START, minute=0))
        available_minutes_today = max(0, (WORK_END - current_time.hour) * 60 - current_time.minute - 1)
        if remaining_minutes <= available_minutes_today:
            current_time += timedelta(minutes=remaining_minutes)
            remaining_minutes = 0
        else:
            remaining_minutes -= available_minutes_today
            current_time = current_time.replace(hour=WORK_START, minute=0) + timedelta(days=1)
            current_time = next_working_day(current_time)
    return current_time

# Function to find gaps in the machine schedule
def find_gaps(machine_schedule):
    gaps = {}
    for machine, tasks in machine_schedule.items():
        tasks = sorted(tasks, key=lambda x: x[0])  # Sort tasks by start time
        gaps[machine] = []
        for i in range(len(tasks) - 1):
            current_end = tasks[i][1]
            next_start = tasks[i + 1][0]
            if next_start > current_end:
                gaps[machine].append((current_end, next_start))  # Record gap
    return gaps

# Main scheduling function with gap optimization
def schedule_production_with_days(data):
    machine_schedule = defaultdict(list)
    machine_last_end = {machine: next_working_day(data['Order Processing Date'].min().replace(hour=WORK_START, minute=0))
                        for machine in data['Machine Number'].unique()}

    for idx, row in data.iterrows():
        component = row['Components']
        product = row['Product Name']
        machine = row['Machine Number']

        # Determine start time
        if "C1" in component and machine == "OutSrc":
            start_time = next_working_day(row['Order Processing Date'].replace(hour=WORK_START, minute=0))
        else:
            prev_component = f"C{int(component[1:]) - 1}"
            same_product_prev = data[(data['Product Name'] == product) & (data['Components'] == prev_component)]
            if not same_product_prev.empty:
                start_time = same_product_prev.iloc[0]['End Time']
            else:
                gaps = find_gaps(machine_schedule)
                for gap_start, gap_end in gaps.get(machine, []):
                    run_time_minutes = row['Run Time (min/1000)'] * row['Quantity Required'] / 1000
                    potential_end_time = adjust_to_working_hours_and_days(gap_start, run_time_minutes)
                    if potential_end_time <= gap_end:  # Check if the task fits in the gap
                        start_time = gap_start
                        break
                else:
                    start_time = max(machine_last_end[machine], next_working_day(row['Order Processing Date'].replace(hour=WORK_START, minute=0)))

            # Adjust start time to working hours
            if start_time.hour < WORK_START:
                start_time = start_time.replace(hour=WORK_START, minute=0)
            elif start_time.hour >= WORK_END or start_time.weekday() in WEEKENDS:
                start_time = next_working_day(start_time.replace(hour=WORK_START, minute=0))

        # Calculate runtime and end time
        run_time_minutes = row['Run Time (min/1000)'] * row['Quantity Required'] / 1000
        end_time = adjust_to_working_hours_and_days(start_time, run_time_minutes)

        # Update machine schedule
        if machine != "OutSrc":
            machine_schedule[machine].append((start_time, end_time, idx))
            machine_schedule[machine] = sorted(machine_schedule[machine], key=lambda x: x[0])
            machine_last_end[machine] = max(machine_last_end[machine], end_time)

        # Assign start and end times to the DataFrame
        data.loc[idx, 'Start Time'] = start_time
        data.loc[idx, 'End Time'] = end_time

    return data

# Function to adjust End Time if it equals {date} 09:00:00
def adjust_end_time(data):
    for idx, row in data.iterrows():
        if row['End Time'].hour == 9 and row['End Time'].minute == 0 and row['End Time'].second == 0:
            data.at[idx, 'End Time'] = row['End Time'] - timedelta(days=1) + timedelta(hours=17)
    return data

# Apply scheduling and adjustments
df = schedule_production_with_days(df)
df = adjust_end_time(df)

# Display the updated dataframe
df[['Product Name', 'Promised Delivery Date', 'Components',
    'Machine Number', 'Operation', 'Quantity Required',
    'Run Time (min/1000)', 'Start Time', 'End Time']]