# -*- coding: utf-8 -*-
"""
Created on Sun Feb  4 11:46:14 2024

@author: WesLauer
@license:  Apache 2.0

The code utilizes the ics.py library, which itself has an Apache 2.0 license.
"""
#from icalendar import Calendar, Event, vCalAddress, vText
from datetime import datetime
import re
import pandas as pd
from dateutil.relativedelta import relativedelta
from ics import Calendar, Event, Organizer
#from zoneinfo import ZoneInfo
from backports.zoneinfo import ZoneInfo
import calendar
import tkinter as tk
from tkinter import filedialog, messagebox

def create_event_ics(title, start_time, end_time, location, organizer, description, calendar_instance):
    e = Event()
    e.name = title
    e.begin = start_time
    e.end = end_time
    e.location = location
    organizer_obj = Organizer(email = 'madeup@mail.com', common_name=organizer)
    e.organizer = organizer_obj
    calendar_instance.events.add(e)
    
# Create a calendar instance
#cal = Calendar()


root = tk.Tk()
root.withdraw()
file_path = filedialog.askopenfilename(title="Select the Excel spreadsheet that includes course schedules")

# Read the Excel file
#filename = '23FQ CLSS Scheduling 022723.xlsx'
filename = file_path
df = pd.read_excel(filename, sheet_name = 'Master Schedule')
df = df.fillna(' ')
df_courses = pd.read_excel(filename, sheet_name = 'Course Schedules to Check')
df_dates = pd.read_excel(filename, sheet_name = 'QuarterStartEnd')

start_of_quarter = datetime.date(df_dates['Start'][0])
end_of_quarter = datetime.date(df_dates['End'][0])
holidays = pd.to_datetime(df_dates['Holidays']).dt.date

courses_to_check = []
for i in range(len(df_courses.axes[1])):
   courses_to_check.append((list(df_courses.iloc[:,i].dropna())))


# Get the schedules from the "Meeting Pattern" column
schedules = df['Meeting Pattern'].tolist()
classtitles = df['Course Title'].tolist()
courses = df['Course'].tolist()
sections = df['Section'].tolist()
meeting_spaces = df['Meeting Space'].tolist()
instructors = df['Instructor'].tolist()

day_map = {
    'M': 0,
    'T': 1,
    'W': 2,
    'Th': 3,
    'F': 4,
    'S': 5,
    'Su': 6
}

# Create events and add them to the calendar instance.
c = Calendar()

for schedule, classtitle, course, meeting_space, instructor, section in zip(schedules, classtitles, courses, meeting_spaces, instructors, sections):
    # Split the schedule into separate events
    events = schedule.split(';')
    for event in events:
        # Parse the days of the week, start time, and end time
        match = re.match(r'([A-Za-z]+) (\d+am|\d+:\d+am|\d+pm|\d+:\d+pm)-(\d+am|\d+:\d+am|\d+pm|\d+:\d+pm)', event.strip())
        days_of_week, start_time, end_time = match.groups()
              
        # Convert the times to 24-hour format
        start_time = datetime.strptime(start_time, "%I:%M%p").time() if ':' in start_time else datetime.strptime(start_time, "%I%p").time()
        end_time = datetime.strptime(end_time, "%I:%M%p").time() if ':' in end_time else datetime.strptime(end_time, "%I%p").time()
        #start_date = datetime.strptime(start_of_quarter, "%m/%d/%y").date()
        start_date = start_of_quarter
        days_of_week = [day_map[day] for day in re.findall(r'(Th|Su|M|T|W|F|S)', days_of_week)]
        
        # Remove duplicates
        days_of_week_integers = list(dict.fromkeys(days_of_week))        
        
        # Add class identifier
        title = course + ' S' + str(section) + ' ' + classtitle  

        organizer = instructor
        
        # Create the event each day of the week
        
        while start_date <= end_of_quarter:
            for day in days_of_week_integers:
                #find the date of the day
                
                next_date = start_date + relativedelta(weekday = day)                    
                start_datetime = datetime.combine(next_date, start_time)
                start_datetime = start_datetime.replace(tzinfo=ZoneInfo('US/Pacific'))
                
                end_datetime = datetime.combine(next_date, end_time)
                end_datetime = end_datetime.replace(tzinfo=ZoneInfo('US/Pacific'))
                
                if start_datetime.date() <= end_of_quarter:
                    if start_datetime.date() not in list(holidays):
                        create_event_ics(title, start_datetime, end_datetime, meeting_space, organizer, classtitle, c)
                        print(title+": "+str(start_datetime))
            start_date = start_date + relativedelta(days=7)
                
# Save calendar to file.
f = filedialog.asksaveasfile(title="Specify output .ics file",mode='w', defaultextension=".ics")
print(f)
if f is None: # asksaveasfile return `None` if dialog closed with "cancel".
    f = 'course_schedule.ics'
else:
    f = f.name
open(f,'w').writelines(c)    


f = filedialog.asksaveasfile(title="Specify course conflicts text file",mode='w', defaultextension=".txt")
print(f)
def check_for_conflicts(test_schedule,master_calendar):
    outputmessage = ''
    student_c = Calendar()
    for course_id in test_schedule:
        for event in list(master_calendar.events):
            if course_id in event.name:
                student_c.events.add(event)
    
    student_courses = list(student_c.timeline)
    conflict = False
    for j in range(len(student_courses)):
        test_course = student_courses[j]
        other_courses = [x for i,x in enumerate(student_courses) if i!=j]    
        for other_course in other_courses:
            does_it_intersect = test_course.intersects(other_course)
            if does_it_intersect:
                if conflict == False:
                    conflict = True
                    outputmessage += 'Conflict present for course schedule ' + str(test_schedule) + '\n\n'
                message = "intersects"
                conflict = True
                day_of_conflict = calendar.day_name[test_course.begin.weekday()]
                time_of_conflict = str(test_course.begin.time())
                outputmessage += str(test_course.name) + ' '+ message + ' ' + str(other_course.name) + ' on ' + day_of_conflict + ' at ' + time_of_conflict + '\n'
    if conflict == False:
        outputmessage = 'No conflicts for course schedule ' + str(test_schedule) + '\n'
    return outputmessage

message_to_write = ''
message_to_write = 'Courses Analysed: ' + str(courses_to_check)
for test_schedule in courses_to_check: 
    message_to_write += '\n******************************\n\n'
    message_to_write += (check_for_conflicts(test_schedule, c))

print(message_to_write)
with open(f.name, 'w') as f:
    f.write(message_to_write)
messagebox.showinfo(title = 'results of course conflict analysis', message=message_to_write)     
    
        
