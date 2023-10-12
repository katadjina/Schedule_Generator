import datetime
import openpyxl
from itertools import cycle

def input_employee_list(shift_name):
    employee_list = []
    print(f"Enter names for the {shift_name} shift. Enter 'done' when finished.")
    
    while True:
        employee_name = input(f"Enter employee for {shift_name}: ").strip()
        if employee_name.lower() == 'done':
            break
        employee_list.append(employee_name)
    
    return employee_list

def input_date(prompt):
    while True:
        try:
            year = int(input(f"Enter {prompt} year (e.g. 2023): "))
            month = int(input(f"Enter {prompt} month (1-12): "))
            day = int(input(f"Enter {prompt} day (1-31): "))

            return datetime.datetime(year, month, day)
        except ValueError:
            print("Invalid date. Please enter a valid date.")


def generate_schedule(employee_morning, employee_afternoon, start_date, end_date):
    morning_cycle = cycle(employee_morning)
    afternoon_cycle = cycle(employee_afternoon)
    
    schedule = []

    # Adjust the start_date to the next Monday if it's not already a Monday
    while start_date.weekday() != 0:
        start_date += datetime.timedelta(days=1)

    current_date = start_date
    while current_date + datetime.timedelta(days=4) <= end_date:  # Ensure we have a full week left
        # The morning and afternoon employees remain the same for a week
        morning_person = next(morning_cycle)
        afternoon_person = next(afternoon_cycle)

        # Append the weekly block to the schedule
        schedule.append((current_date, current_date + datetime.timedelta(days=4), morning_person, afternoon_person))
        
        # Jump to the next Monday
        current_date += datetime.timedelta(days=7)

    return schedule

def write_to_excel(schedule, filename):
    workbook = openpyxl.Workbook()
    sheet = workbook.active
    sheet.append(['Start Date', 'End Date', 'Employee Morning', 'Employee Afternoon'])

    for start_day, end_day, morning, afternoon in schedule:
        sheet.append([start_day.strftime('%Y-%m-%d'), end_day.strftime('%Y-%m-%d'), morning, afternoon])

    workbook.save(filename)

if __name__ == "__main__":
    employee_morning = input_employee_list("morning")
    employee_afternoon = input_employee_list("afternoon")
    
    print("\nEnter the start date:")
    start_date = input_date("start")

    print("\nEnter the end date:")
    end_date = input_date("end")

    schedule = generate_schedule(employee_morning, employee_afternoon, start_date, end_date)
    write_to_excel(schedule, "Employee_Schedule.xlsx")
    
    print("Schedule saved to Employee_Schedule.xlsx!")

