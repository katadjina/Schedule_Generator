# Schedule Generator

## This script automates the task of creating weekly schedules for employees. It generates a schedule based on the availability of employees for morning and afternoon shifts.

Features
Interactive input for employee names for both morning and afternoon shifts
Flexible start and end dates for the schedule
Generates an Excel sheet with the weekly schedule, starting from Monday and ending on Friday
Ensures that the shifts rotate between the available employees in the order they're entered

Requirements

Python 3
openpyxl library for handling Excel operations
To install the required library, run:


<pip install openpyxl>
  
Usage
Clone the repository or download the script.
Navigate to the directory containing the script using a terminal or command prompt.
Run the script using:
bash
Copy code
python script_name.py
Replace script_name.py with the name of the Python script.

Follow the interactive prompts to input the names of employees for morning and afternoon shifts, and specify the start and end dates for the schedule.
The script will generate an Excel file named Employee_Schedule.xlsx in the current directory with the weekly schedule.
