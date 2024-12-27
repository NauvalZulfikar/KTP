import pandas as pd
#import time
import time as tm
import xlwings as xw
from datetime import datetime, timedelta, date,time
import plotly.express as px
from IPython.display import display, clear_output
import os
import plotly.graph_objects as go
import altair as alt
import subprocess
import openpyxl
import sqlalchemy
import urllib.parse
from sqlalchemy import create_engine, MetaData, Table, Column, Integer, String, DateTime
import psycopg2
from psycopg2 import sql
from psycopg2.extras import execute_values
import sys
import queue
from multiprocessing import Queue, current_process, get_context
import argparse
import traceback
import logging
from openpyxl import load_workbook
from openpyxl.styles import Font

output_queue = queue.Queue()

# Initialize machine status dictionary
machine_status = {}
O1 = 0

interrupted = False
def read_excel(file_path, sheet_name):
    return pd.read_excel(file_path, sheet_name=sheet_name)

def write_excel(df, file_path, sheet_name):
    print(df)
    with pd.ExcelWriter(file_path, mode='a', engine='openpyxl', if_sheet_exists='replace') as writer:
        df.to_excel(writer, sheet_name=sheet_name, index=False)

def update_runtime(run_time, file_path):
    file_path="RunTime.xlsx"
    runtime_df = pd.DataFrame({'Run_time': [run_time]})
    write_excel(runtime_df, file_path, 'RunTime')


def fetch_data(file_path):
    try:
        addln_df = pd.read_excel(file_path, sheet_name='Addln')
        prodet_df = pd.read_excel(file_path, sheet_name='prodet')

        if not addln_df.empty:
            prodet_df = pd.concat([prodet_df, addln_df], ignore_index=True)
            addln_df = pd.DataFrame()
            # Use the openpyxl engine and retain existing sheets
            with pd.ExcelWriter(file_path, engine='openpyxl', mode='a', if_sheet_exists='replace') as writer:
                # Save updated sheets back to the file
                prodet_df.to_excel(writer, sheet_name='prodet', index=False)
                addln_df.to_excel(writer, sheet_name='Addln', index=False)

        return prodet_df
    except Exception as e:
        print(f"Error fetching data: {e}")
        return pd.DataFrame()

def fetch_latest_completed_time(file_path, sheet_name='prodet'):
    try:
        # Read the Excel file
        prodet_df = pd.read_excel(file_path, sheet_name=sheet_name)
        
        # Filter the DataFrame for rows where 'Status' is 'Completed'
        completed_df = prodet_df[prodet_df['Status'] == 'Completed']
        
        # Find the maximum 'End Time' in the filtered DataFrame
        if not completed_df.empty:
            #latest_end_time_str = completed_df['End Time'].max()
            # Convert the 'End Time' column to datetime format if it's not already
            completed_df['End Time'] = pd.to_datetime(completed_df['End Time'], format='%Y-%m-%d %H:%M:%S')
            
            # Get the row with the latest 'End Time'
            latest_row = completed_df.loc[completed_df['End Time'].idxmax()]
            # Extract the 'End Time' and 'Machine Number'
            latest_end_time = latest_row['End Time']
            machine_number = latest_row['Machine Number']
            # Convert the end time to a datetime object if it is a string
            #if isinstance(latest_end_time_str, str):
                #return datetime.strptime(latest_end_time_str, '%Y-%m-%d %H:%M:%S')
            #return latest_end_time_str
            # Return the end time and machine number
            print("Latest End Time:", latest_end_time)
            print("Machine Number:", machine_number)
            
            return latest_end_time, machine_number
        else:
            print("No completed products found")
            return None, None  # If returning from a function
    except Exception as e:
        print(f"Error fetching latest completed time: {e}")
        return None, None  # If returning from a function

def calculate_remaining_time(product_name, components_df):
    remaining_components = components_df[(components_df['Product Name'] == product_name) & 
                                         (components_df['Status'] != 'Completed')]
    total_remaining_time = 0
    for index, row in remaining_components.iterrows():
        run_time_per_1000 = row['Run Time (min/1000)']
        quantity = row['Quantity Required']
        cycle_time = (run_time_per_1000 * quantity) / 1000
        total_remaining_time += cycle_time * 60  # Convert to seconds
    return total_remaining_time

work_start = 9
work_end = 17
work_days = [0, 1, 2, 3, 4]  # Monday to Friday

def next_working_day(current_time):
    while current_time.weekday() not in work_days:
        current_time += timedelta(days=1)
    return current_time

def calculate_end_time(start_time, duration_minutes):
    current_time = start_time
    remaining_minutes = duration_minutes

    while remaining_minutes > 0:
        # If the current time is outside working hours, adjust to next working start time
        if current_time.hour < work_start:
            current_time = current_time.replace(hour=work_start, minute=0, second=0)
        elif current_time.hour >= work_end or current_time.weekday() not in work_days:
            current_time += timedelta(days=1)
            current_time = current_time.replace(hour=work_start, minute=0, second=0)
            current_time = next_working_day(current_time)

        end_of_day = current_time.replace(hour=work_end, minute=0, second=0)
        available_minutes = (end_of_day - current_time).total_seconds() // 60

        if remaining_minutes <= available_minutes:
            current_time += timedelta(minutes=remaining_minutes)
            remaining_minutes = 0
        else:
            remaining_minutes -= available_minutes
            current_time = end_of_day + timedelta(days=1)
            current_time = current_time.replace(hour=work_start, minute=0, second=0)
            current_time = next_working_day(current_time)

    return current_time

def allocate_machines(outsource_df, components_df, machines_df, Similarity_df,input_data,file_path):
    global O1
    Prev_MachNumber=None
    frststep=0
    flag_update=0
    OrderProcessingDate = components_df['Order Processing Date'].min()
    machine_status = {machine: 0 for machine in machines_df['Machines'].tolist()}
    #last_processed = {machine: (None, None, None, None) for machine in machines_df['Machines'].tolist()}  # To store last processed (component, machine, operation)
    last_processed = {machine: (None, None, None, None, None) for machine in machines_df['Machines'].tolist()}  # To store last processed (component, machine, operation)
    # Outcycle_time = 20
    CurrentOutSrcPN = ''
    simulated_time = OrderProcessingDate.replace(hour=9, minute=0, second=0, microsecond=0)  # Start time at 9 am
    #simulated_time = datetime.now().replace(hour=9, minute=0, second=0, microsecond=0)  # Start time at 9 am
    #simulated_time = datetime.now()
    ReadyTime = OrderProcessingDate.replace(hour=9, minute=0, second=0, microsecond=0)  # Start time at 9 am
    animation_data = []
    global interrupted
    #print(components_df)

    if input_data == "Start":

        # Initialize O1 and CurrentPN
        O1 = 0
        CurrentPN = None

        # Check for any 'Outsrc' component that is 'InProgress'
        in_progress_outsrc = components_df[(components_df['Status'] == 'InProgress') & (components_df['Process Type'] == 'Outsource')]

        print(in_progress_outsrc)
        # If there are any such components, set O1 and CurrentPN
        if not in_progress_outsrc.empty:
            O1 = 1
            CurrentPN = in_progress_outsrc.iloc[0]['Product Name']

        simulated_time,Prev_MachNumber=fetch_latest_completed_time(file_path)
        end_time=simulated_time
        print(f"Start Simulated Time: {simulated_time}")

    while True:
        if interrupted:
            break
        if simulated_time.hour >= 17:
            #print("Stopping the bot as it is 17:00")
            #break
            
            simulated_time = OrderProcessingDate.replace(hour=9, minute=0, second=0, microsecond=0)  # Start time at 9 am
            simulated_date= simulated_time.strftime('%A')
            if simulated_date=="Friday":
                simulated_time += timedelta(days=3)
            elif simulated_date=="Saturday":
                simulated_time += timedelta(days=2)
            else:
                simulated_time += timedelta(days=1)
            print("After 17hrs: {simulated_time}")
        

        
        #if 12 <= simulated_time.hour < 13:
         #   print("Pausing the bot for lunch break (12:00 PM to 1:00 PM)")
          #  simulated_time += timedelta(hours=1)
           # continue
        
        components_df=fetch_data(file_path)  
        print(components_df)
        remaining_components = components_df[(components_df['Status'] != 'Completed') & (components_df['Status'] != 'Late')]
        #print(remaining_components)
        if remaining_components.empty:
            break
        
        # Get today's date as a datetime object
        today = date.today()

        # Convert today's date to string in the desired format
        today_date_str = today.strftime("%d-%m-%Y")

        # Convert today_date_str back to a datetime object
        today_date = datetime.strptime(today_date_str, "%d-%m-%Y")

        df=components_df

        # Convert 'Promised Delivery Date' column to datetime if it's not already
        df['Promised Delivery Date'] = pd.to_datetime(df['Promised Delivery Date'])

        # Calculate remaining days
        df["Remaining_days"] = (df['Promised Delivery Date'] - today_date).dt.days

        # Convert remaining days to minutes and adjust by subtracting 960 minutes per day
        df["Time_Needed"] = df["Remaining_days"] * 24 * 60 - df["Remaining_days"] * 960

        # Use vectorized operations for conditional logic
        df['Run_time'] = df.apply(
        lambda row: row['Run Time (min/1000)'] if row['Process Type'] == "Outsource" else (row['Run Time (min/1000)'] * row['Quantity Required']) / 1000,
        axis=1
        )

        product_times = {}

        for product in df['Product Name'].unique():
            product_df = df[df['Product Name'] == product]
            
            total_time = product_df['Run_time'].sum()
            remaining_time = product_df[product_df['Status'] != 'Completed']['Run_time'].sum()
            remaining_days = product_df['Time_Needed'].mean()
            Diff = remaining_days - remaining_time
            product_times[product] = {
                'Total Time': total_time,
                'Remaining Time': remaining_time,
                'Time_Needed': remaining_days,
                'Time Left': Diff
            }

        time_df = pd.DataFrame(product_times).T.reset_index().rename(columns={'index': 'Product Name'})
        #print(time_df)
        work_start = time(9, 0)
        work_end = time(17, 0)

        # Current time for the example
        current_time=simulated_time
        #current_time = datetime.now()
        remaining_time_minutes = 0
        dates = []  # List to store the calculated dates

        for i, r in time_df.iterrows():
            remaining_time_minutes_r = r['Remaining Time']
            #print(i)
            if i > 0:
                remaining_time_minutes += remaining_time_minutes_r
            else:
                remaining_time_minutes = remaining_time_minutes_r

            # Calculate remaining working minutes in the current day
            if current_time.time() < work_start:
                remaining_minutes_today = 8 * 60  # Full working day available
            elif current_time.time() > work_end:
                remaining_minutes_today = 0  # No working time left today
            else:
                remaining_minutes_today = ((work_end.hour - current_time.hour) * 60 - current_time.minute)

            if remaining_time_minutes <= remaining_minutes_today:
                date_val = today_date.strftime("%d-%m-%Y")
            else:
                # Calculate how many full working days are needed
                full_days_needed = remaining_time_minutes_r // (8 * 60)
                remaining_minutes = remaining_time_minutes_r % (8 * 60)
                completion_date = today_date
                
                # Skip weekends
                while full_days_needed > 0:
                    completion_date += timedelta(days=1)
                    if completion_date.weekday() < 5:  # Only count weekdays
                        full_days_needed -= 1
                
                if remaining_minutes > 0:
                    completion_date += timedelta(days=1)
                    while completion_date.weekday() >= 5:  # Skip weekends
                        completion_date += timedelta(days=1)
                
                date_val = completion_date.strftime("%d-%m-%Y")
                
            dates.append(date_val)  # Append the calculated date to the list
            print(date_val)

        time_df['Date'] = dates  # Add the list as a new column to time_df

        # Sort time_df by 'Time Left' in ascending order
        time_df_sorted = time_df.sort_values(by='Time Left', ascending=True)

        # Get the first product name from the sorted DataFrame
        first_product_name = time_df_sorted.iloc[0]['Product Name']

        # Print the sorted DataFrame
        print(time_df_sorted)

        # Print the first product name
        print("First Product Name:", first_product_name)

        #if flag_update==0:
            #update_timedata(time_df_sorted)
            #flag_update=1
        # Get the first product name from the sorted DataFrame
        #first_product_name = time_df_sorted.iloc[0]['Product Name']

        #print("First Product Name:", first_product_name)
        for ir, rr in time_df_sorted.iterrows():
            first_product_name=rr['Product Name']
            remaining_components = components_df[(components_df['Status'] != 'Completed') & (components_df['Status'] != 'Late')]
            if remaining_components.empty:
               break
            remaining_components = remaining_components[remaining_components['Product Name'] == first_product_name]
            print(remaining_components)
            #remaining_components = components_df[components_df['Status'] != 'Completed']
            #print(remaining_components)
            

            for index, row in remaining_components.iterrows():
                #print(f"{index}Index")
                component = row['Components']
                cycle_time_r = row['Run Time (min/1000)']
                Qnty = row['Quantity Required']
                cycle_time = (cycle_time_r * Qnty) / 1000  # in minutes
                cycle_time_seconds = cycle_time * 60  # convert cycle_time to seconds
                machine_number = row['Machine Number']
                status = row['Status']
                ProductNames = row['Product Name']
                ProcessType = row['Process Type']
                PromisedDeliveryDate = row['Promised Delivery Date']
                operation = row['Operation']
                Outcycle_time_1 = row['Run Time (min/1000)']
                CurrentTimeStrt=datetime.now()
            
                
                if simulated_time.hour >=17:
                    #print("Stopping the bot as it is 17:00")
                    #break
                    
                    simulated_time = OrderProcessingDate.replace(hour=9, minute=0, second=0, microsecond=0)  # Start time at 9 am
                    simulated_date= simulated_time.strftime('%A')
                    if simulated_date=="Friday":
                        simulated_time += timedelta(days=3)
                    elif simulated_date=="Saturday":
                        simulated_time += timedelta(days=2)
                    else:
                        simulated_time += timedelta(days=1)
                    print("After 17hrs: {simulated_time}")
                if O1 == 1 and CurrentOutSrcPN == ProductNames:
                    continue

                if (O1 == 1 and ProductNames == CurrentPN) or (status == "InProgress" and ProcessType=="Outsource"):
                    CurrIndex = index
                    
                    for ind, ro in outsource_df.iterrows():
                        ProdName=ro['Product']
                        CompoName=ro['Components']
                        if ProdName==ProductNames:
                            Outcycle_time=ro['Outsource Time']
                            break
                    Outcycle_time=20   
                    print(Outcycle_time)
                    
                    OutsrcSrt_Time = pd.to_datetime(components_df.loc[CurrIndex, 'Start Time'], errors='coerce')
                    print(OutsrcSrt_Time)

                    if OutsrcSrt_Time!=None:
                        #simulated_time += timedelta(minutes=Outcycle_time_1)
                        current_Ptime=simulated_time
                        #print(f"{current_Ptime}, {OutsrcSrt_Time}, {Outcycle_time}")
                        #print(f"{CurrIndex} CurrentPtime(Sim) diff OutsrcTime")
                        l=current_Ptime-OutsrcSrt_Time
                        d=l.total_seconds()
                        #print(f"{d}, {CurrIndex} CurrentPtime(Sim) diff OutsrcTime")
                
                        if l.total_seconds() >= Outcycle_time:
                            #print(f"{current_Ptime}, {OutsrcSrt_Time}, {Outcycle_time}")
                            O1 = 0
                            #print(f"OutSrcStartTime: {OutsrcSrt_Time}")
                            #OutsrcSrt_Time+= timedelta(minutes=Outcycle_time_1)
                            #print(f"OutSrcStartTime: {OutsrcSrt_Time}")
                            #print(f"Outcycle_time_1: {Outcycle_time_1}")
                            
                            #O1End_time = OutsrcSrt_Time
                            Ref_time=OutsrcSrt_Time-simulated_time
                            print(Ref_time)
                            last_comp, last_machine, last_op, last_prod, last_endtime = last_processed[Prev_MachNumber]
                            
                            O1End_time=calculate_end_time(OutsrcSrt_Time, Outcycle_time_1)
                            if last_endtime==None:
                                Final_EndTime=O1End_time
                            else:
                                if O1End_time > last_endtime:
                                    Final_EndTime=O1End_time
                                else:
                                    Final_EndTime=last_endtime
                            components_df.loc[CurrIndex, 'End Time'] = Final_EndTime.strftime("%Y-%m-%d %H:%M:%S")  # Update only time
                            components_df.loc[CurrIndex, 'Status'] = 'Completed'
                            #update_runtime(Outcycle_time_1)
                            #simulated_time=OutsrcSrt_Time
                            setup_time=30
                            simulated_time=Final_EndTime + timedelta(seconds=setup_time)
                            #simulated_time=Final_EndTime
                            # Update the last processed product details
                            last_processed[machine_number] = (component, machine_number, operation, ProductNames, end_time)
                            Prev_MachNumber=machine_number
                            CurrentOutSrcPN = ''
                            CurrIndex = ''
                            write_excel(components_df, file_path, 'prodet')
                            continue
                            #break

                    if ProductNames == CurrentPN:
                        #CurrentTimEnd=datetime.now()
                        #diff= CurrentTimEnd-CurrentTimeStrt
                        #simulated_time = simulated_time + diff
                        continue

                if status == 'InProgress' and ProcessType == 'Outsource':
                    CurrentOutSrcPN = ProductNames
                    write_excel(components_df, file_path, 'prodet')
                    #CurrentTimEnd=datetime.now()
                    #diff= CurrentTimEnd-CurrentTimeStrt
                    #simulated_time = simulated_time + diff
                    continue

                if status != 'Completed' or status!='Late':
                    if ProcessType == 'Outsource':
                        O1 = 1
                        OutsrcSrt_Time = simulated_time
                        components_df.loc[index, 'Start Time'] = OutsrcSrt_Time.strftime("%Y-%m-%d %H:%M:%S")  # Update only time
                        components_df.loc[index, 'Status'] = 'InProgress'
                        CurrIndex = index
                        CurrentPN = ProductNames
                        write_excel(components_df, file_path, 'prodet')
                        continue
                    else:
                        if machine_status[machine_number] == 0:
                            # Calculate wait time and update it
                            wait_time = simulated_time - ReadyTime
                            #print(wait_time)
                            wait_time_str = f"{wait_time.seconds // 3600:02}:{(wait_time.seconds // 60) % 60:02}:{wait_time.seconds % 60:02}"
                            #wait_time_str = f"{wait_time.seconds // 3600:02}:{(wait_time.seconds // 60) % 60:02}"
                            components_df.loc[index, 'Wait Time'] = wait_time_str  # Update only time
                            
                            machine_status[machine_number] = 1
                            start_time = simulated_time
                            components_df.loc[index, 'Start Time'] = start_time.strftime("%Y-%m-%d %H:%M:%S")  # Update only time
                            components_df.loc[index, 'Status'] = 'InProgress'
                            write_excel(components_df, file_path, 'prodet')
                            #start_clock(machine_number)
                            #time.sleep(cycle_time)  # Convert cycle_time to seconds

                            # Check similarity status and determine setup time
                            similarity_status=row['SetupTimeCheck']
                            if similarity_status == 1:
                                setup_time = 0  # No setup time needed
                            else:
                                setup_time = row['Setup time (seconds)']  # Setup time is 5 minutes
                            

                            components_df.loc[index, 'Final Setup Time'] = setup_time
                            print(setup_time)
                            #Run_time_r=setup_time + cycle_time
                            Run_time_r=cycle_time
                            tm.sleep(5)  # Wait for setup time and cycle time
                            end_time=calculate_end_time(start_time, cycle_time)
                            # Update the last processed product details
                            #last_processed[machine_number] = (component, machine_number, operation, ProductNames)
                            last_processed[machine_number] = (component, machine_number, operation, ProductNames, end_time)
                            simulated_time += timedelta(minutes=Run_time_r)  # Update simulated time
                            machine_status[machine_number] = 0
                            #end_time = simulated_time
                            
                            setup_time=30
                            simulated_time=end_time + timedelta(seconds=setup_time)
                            components_df.loc[index, 'End Time'] = end_time.strftime("%Y-%m-%d %H:%M:%S")  # Update only time
                            if end_time>PromisedDeliveryDate:
                                components_df.loc[index, 'Status'] = 'Late'
                                N_PDDate=PromisedDeliveryDate.replace(hour=9, minute=0, second=0, microsecond=0)
                                Delay_Time=end_time-N_PDDate
                                # Calculate delay in days and hours

                                delay_days = Delay_Time.days
                                delay_hours = Delay_Time.seconds//3600
                                
                                # Store the values in their respective columns
                                components_df.loc[index, 'Delay Days'] = delay_days
                                components_df.loc[index, 'Delay Hours'] = delay_hours
                            else:
                                components_df.loc[index, 'Status'] = 'Completed'
                                components_df.loc[index, 'Delay Days'] = 0
                                components_df.loc[index, 'Delay Hours'] = 0


                            #update_runtime(Run_time_r)
                            write_excel(components_df, file_path, 'prodet')
                            #stop_clock(machine_number)
                            Prev_MachNumber=machine_number
                write_excel(components_df, file_path, 'prodet')
                tm.sleep(1)

                # Collect animation data
                current_time = simulated_time.strftime("%H:%M:%S")
                for i, row in components_df.iterrows():
                    animation_data.append(dict(
                        Time=current_time,
                        Product=row['Product Name'],
                        Component=row['Components'],
                        Machine=row['Machine Number'],
                        Status=row['Status']
                    ))
                #create_gantt_chart(components_df)
                #CurrentTimEnd=datetime.now() - check
                #diff= CurrentTimEnd-CurrentTimeStrt
                #simulated_time = simulated_time + diff
    return time_df_sorted 

def main():
        
      
        file_path = r'Product Details_v2.xlsx'
        print(file_path)
        input_data = sys.stdin.read().strip()
        #input_data="Initial"
        print(input_data)
        if input_data == "Initial":
            
            print("Initialization process started.")
            components_df = pd.read_excel(file_path, sheet_name='P')
            outsource_df = pd.read_excel(file_path, sheet_name='Outsource Time')
            machines_df = pd.read_excel(file_path, sheet_name='Machines')
            Similarity_df = pd.read_excel(file_path, sheet_name='Similarity')
            # Convert DataFrame to dictionary
            similarity_lookup = Similarity_df.set_index('Machine').T.to_dict()
            for index, row in components_df.iterrows():
                product = row["Product Name"]
                machine = row["Machine Number"]
                
                if machine == "OutSrc":
                    components_df.loc[index, 'SetupTimeCheck'] = 1
                else:
                    if machine in similarity_lookup and product in similarity_lookup[machine]:
                        if similarity_lookup[machine][product] == 1:
                            components_df.loc[index, 'SetupTimeCheck'] = 1
                        else:
                            components_df.loc[index, 'SetupTimeCheck'] = 0
                    else:
                        components_df.loc[index, 'SetupTimeCheck'] = 0
            print(components_df)
                

            write_excel(components_df, file_path, 'prodet')

            time_df_sorted=allocate_machines(outsource_df, components_df, machines_df, Similarity_df,input_data,file_path)
        elif input_data == "Start":
            print("Start process initiated.")
            outsource_df = pd.read_excel(file_path, sheet_name='Outsource Time')
            machines_df = pd.read_excel(file_path, sheet_name='Machines')
            Similarity_df = pd.read_excel(file_path, sheet_name='Similarity')
            components_df=fetch_data(file_path)
            # Convert DataFrame to dictionary
            similarity_lookup = Similarity_df.set_index('Machine').T.to_dict()
            for index, row in components_df.iterrows():
                product = row["Product Name"]
                machine = row["Machine Number"]
                
                if machine in similarity_lookup and product in similarity_lookup[machine]:
                    if similarity_lookup[machine][product] == 1:
                        components_df.loc[index, 'SetupTimeCheck'] = 0
                else:
                    components_df.loc[index, 'SetupTimeCheck'] = 1

            #print(components_df)
            #write_excel(components_df, file_path, 'prodet')
            time_df_sorted=allocate_machines(outsource_df, components_df, machines_df, Similarity_df,input_data,file_path)
        else:
            print(f"Unknown command: {input_data}")
    

        components_df=fetch_data(file_path)
        #print(components_df)
        # Convert 'Start Time' and 'End Time' columns to datetime
        components_df['Start Time'] = pd.to_datetime(components_df['Start Time'])
        components_df['End Time'] = pd.to_datetime(components_df['End Time'])

        # Calculate the duration (difference) between start and end times
        components_df['Time Diff_days'] = components_df['End Time'] - components_df['Start Time']

        print(components_df)
        # Fixed total time (start time in this case)
        total_time = datetime.now().replace(hour=9, minute=0, second=0, microsecond=0)

        # Calculate Idle Time
        components_df['Idle Time'] = total_time - components_df['Time Diff_days']

        # Format Idle Time to %H:%M:%S
        components_df['Idle Time'] = components_df['Idle Time'].dt.strftime('%H:%M')
        
        print(components_df)
        #update_excel(components_df,connection)
        # Format 'Time Diff' to %H:%M:%S
        #components_df['Time Diff'] = components_df['Time Diff_days'].dt.total_seconds().apply(lambda x: pd.Timedelta(seconds=x)).dt.floor('s').astype(str)
        #components_df['Time Diff'] = components_df['Time Diff'].apply(lambda x: str(x).split(' ')[-1])
        # Step 1: Convert the time difference from timedelta to total seconds
        components_df['Time Diff'] = components_df['Time Diff_days'].dt.total_seconds()

        # Step 2: Convert total seconds to hours and minutes format '%H:%M'
        components_df['Time Diff'] = components_df['Time Diff'].apply(lambda x: pd.to_datetime(x, unit='s').strftime('%H:%M:%S'))

        components_df['Promised Delivery Date'] = components_df['Promised Delivery Date'].dt.strftime('%Y-%m-%d %H:%M:%S')
        components_df['Order Processing Date'] = components_df['Order Processing Date'].dt.strftime('%Y-%m-%d %H:%M:%S')
        components_df['Start Time'] = components_df['Start Time'].dt.strftime('%Y-%m-%d %H:%M:%S')
        components_df['End Time'] = components_df['End Time'].dt.strftime('%Y-%m-%d %H:%M:%S')
        
        

        write_excel(components_df, file_path, 'prodet')
        write_excel(time_df_sorted, file_path, 'Product Times')
        write_excel(outsource_df, file_path, 'Outsource Time')
        write_excel(machines_df, file_path, 'Machines')
        write_excel(Similarity_df, file_path, 'Similarity')
       
        wb = load_workbook(file_path)
        ws = wb['prodet']  # Assuming you want to apply formatting on this sheet

        # Define red font style for highlighting
        red_font = Font(color="FF0000", bold=True)

        # Iterate over the rows in the DataFrame
        components_df = pd.read_excel(file_path, sheet_name='prodet')

        # Read the sheet where "Status" is located, and apply the red font if status is "Late"
        for row_idx, row in components_df.iterrows():
            status = row['Status']  # Assuming 'Status' column holds the status
            if status == "Late":
                # Highlight entire row (columns A to Z, or adjust as needed)
                excel_row = row_idx + 2  # Adding 2 because row indices in openpyxl are 1-based, and row 1 is usually the header
                for col in range(1, 27):  # Columns A (1) to Z (26)
                    ws.cell(row=excel_row, column=col).font = red_font

        # Save the workbook with the applied font style
        wb.save(file_path)


if __name__ == "__main__":
    main()
