import pandas as pd
import random
from datetime import datetime, timedelta

def daterange(start_date, end_date):
    for n in range(int ((end_date - start_date).days)):
        yield start_date + timedelta(n)
        
def export_schedule(schedule):
    schedule.to_excel('schedule.xlsx', index=False)
    print("Schedule exported successfully to schedule.xlsx")

# Define the schedule frame
frame = pd.DataFrame(columns=['Worker', 'Date', 'Shift', 'Week_number'])
    
# Create a list of workers
workers = ["Tomaž Čuček", "Danijel Mršič", "Uroš Petrič", "Franc Primožič", "Samo Vesel", "Primož Bajec", "Alen Mašić", "Gorazd Udovič", "Simona Zadnik", "Bernard Kokol", "Mojca Česnik", "Nina Eder", "Martin Velikanje",  "Urban Urbanc",
"Nejc Cimprič", "Sabrina Jeranče", "Mitja Tekavec",  "Damjan Kokol", "Jagodič Borut", "Albreht Matjaž", "Černivšek Urh", "Aneja Šavli", "Mlakar Bojan", "Nejedly Sebastjan", "Ajster Vid",]

# Define the start and end dates
start_date = "01-02-2023"
end_date = "01-03-2023"

# Convert start and end date to datetime object
start_date = datetime.strptime(start_date, "%d-%m-%Y")
end_date = datetime.strptime(end_date, "%d-%m-%Y")

# Get the list of dates
dates = list(daterange(start_date, end_date))

# Define the shifts
shifts = [['07:00', '19:00'], ['19:00', '07:00']]

# Generate the schedule
worker_shifts = {}
for worker in workers:
    worker_shifts[worker] = {'month_hours': 0, 'week_hours': 0}
    
for date in dates:
    work1_day_shift_workers = random.sample(workers, 2)
    work1_night_shift_workers = random.sample(workers, 2)
    work2_day_shift_workers = random.sample(workers, 2)
    work2_night_shift_workers = random.sample(workers, 2)

    for worker in work1_day_shift_workers:
        if (worker_shifts[worker]['month_hours'] + 12) < 168:
            start_time = '07:00'
            end_time = '19:00'
            schedule = pd.DataFrame({'Worker': worker, 'Date': date, 'Start Time': start_time, 'End Time': end_time,'Shift':'Day','Work':'Basic_patrol', 'Week_number' : [int((date - start_date).days / 7) + 1]}, index=[0])
            frame = pd.concat([frame, schedule], ignore_index=True)        
            worker_shifts[worker]['month_hours'] += 12
            worker_shifts[worker]['week_hours'] += 12  
            if worker_shifts[worker]['week_hours'] == 56:
                worker_shifts[worker]['week_hours'] = 0
    
    for worker in work2_day_shift_workers:
        if (worker_shifts[worker]['month_hours'] + 12) < 168:
            start_time = '07:00'
            end_time = '19:00'
            schedule = pd.DataFrame({'Worker': worker, 'Date': date, 'Start Time': start_time, 'End Time': end_time,'Shift':'Day','Work':'izravnalni', 'Week_number' : [int((date - start_date).days / 7) + 1]}, index=[0])
            frame = pd.concat([frame, schedule], ignore_index=True) 
            worker_shifts[worker]['month_hours'] += 12
            worker_shifts[worker]['week_hours'] += 12  
            if worker_shifts[worker]['week_hours'] == 56:
                worker_shifts[worker]['week_hours'] = 0
    
    for worker in work1_night_shift_workers:
        if (worker_shifts[worker]['month_hours'] + 12) < 168:
            start_time = '07:00'
            end_time = '19:00'
            schedule = pd.DataFrame({'Worker': worker, 'Date': date, 'Start Time': start_time, 'End Time': end_time,'Shift':'Day','Work':'Basic_patrol', 'Week_number' : [int((date - start_date).days / 7) + 1]}, index=[0])
            frame = pd.concat([frame, schedule], ignore_index=True)
            worker_shifts[worker]['month_hours'] += 12
            worker_shifts[worker]['week_hours'] += 12  
            if worker_shifts[worker]['week_hours'] == 56:
                worker_shifts[worker]['week_hours'] = 0
    
    for worker in work2_night_shift_workers:
        if (worker_shifts[worker]['month_hours'] + 12) < 168:
            start_time = '07:00'
            end_time = '19:00'
            schedule = pd.DataFrame({'Worker': worker, 'Date': date, 'Start Time': start_time, 'End Time': end_time,'Shift':'Day','Work':'izravnalni', 'Week_number' : [int((date - start_date).days / 7) + 1]}, index=[0])
            frame = pd.concat([frame, schedule], ignore_index=True)
            worker_shifts[worker]['month_hours'] += 12
            worker_shifts[worker]['week_hours'] += 12  
            if worker_shifts[worker]['week_hours'] == 56:
                worker_shifts[worker]['week_hours'] = 0

# call the function to export the schedule
export_schedule(frame)
for worker, hours in worker_shifts.items():
    print(f'{worker} worked {hours["month_hours"]} hours this month and {hours["week_hours"]} hours this week.')     

# Create an empty DataFrame to store the worker hours data
worker_hours_df = pd.DataFrame(columns=['Worker', 'Month_hours'])

# Iterate over the worker_shifts dictionary and add the worker and month hours to the DataFrame
for worker, hours in worker_shifts.items():
    worker_hours_df = pd.concat([worker_hours_df, pd.DataFrame({'Worker': [worker], 'Month_hours': [hours['month_hours']], 'Avg. week_hours': [hours['month_hours'] / 4]})], ignore_index=True)
    print(f"{worker} worked {hours['month_hours']} hours this month and {hours['month_hours'] / 4} avg. hours per week.")

# Export the DataFrame to an Excel file
worker_hours_df.to_excel("worker_hours.xlsx", sheet_name='Worker_hours', index=False)