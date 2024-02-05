# -*- coding: utf-8 -*-
"""
Created on Sun Feb  4 11:46:14 2024

@author: WesLauer

"""

from icalendar import Calendar, Event
from datetime import datetime
import re
import pandas as pd
#import pytz
from icalendar import vCalAddress, vText


start_of_quarter = input("Enter start date of quarter in mm/dd/yy format or blank for 02/01/24") or "02/01/24" 
end_of_quarter = input("Enter last class day of quarter in mm/dd/yy format or blank for 02/22/24") or "02/22/24" 
end_of_quarter_date = datetime.strptime(end_of_quarter, "%m/%d/%y").date()

# Function to create an event
def create_event(title, start_time, end_time, days, location, organizer, description):
    event = Event()
    event.add('dtstart', start_time)
    event.add('dtend', end_time)
    event.add('summary', title+' '+description)
    #event.add('description', description)
    event.add('location', location)
    #organizer.params['cn'] = vText(instructor)
    event['organizer'] = organizer
    
    #event.add('attendee', organizer)
    #event.add('required', organizer)
    
    # Repeat rule based on days of the week
    repeat_days = ','.join([day[:2].upper() for day in days])
    my_list = repeat_days.split(',')
    event.add('rrule', {'freq': 'weekly', 'byday': my_list, 'until': end_of_quarter_date})
   
    return event

# Create a calendar instance
cal = Calendar()

# Read the Excel file
df = pd.read_excel('23FQ CLSS Scheduling 022723.xlsx')
df = df.fillna(' ')

# Get the schedules from the "Meeting Pattern" column
schedules = df['Meeting Pattern'].tolist()
classtitles = df['Course Title'].tolist()
courses = df['Course'].tolist()
meeting_spaces = df['Meeting Space'].tolist()
instructors = df['Instructor'].tolist()

day_map = {
    'M': 'Mon',
    'T': 'Tue',
    'W': 'Wed',
    'Th': 'Thu',
    'F': 'Fri',
    'S': 'Sat',
    'Su': 'Sun'
}

# Create events and add them to the calendar instance.

for schedule, classtitle, course, meeting_space, instructor in zip(schedules, classtitles, courses, meeting_spaces, instructors):
    # Split the schedule into separate events
    events = schedule.split(';')
    for event in events:
        # Parse the days of the week, start time, and end time
        match = re.match(r'([A-Za-z]+) (\d+am|\d+:\d+am|\d+pm|\d+:\d+pm)-(\d+am|\d+:\d+am|\d+pm|\d+:\d+pm)', event.strip())
        days_of_week, start_time, end_time = match.groups()
        
        
        # Convert the times to 24-hour format
        start_time = datetime.strptime(start_time, "%I:%M%p").time() if ':' in start_time else datetime.strptime(start_time, "%I%p").time()
        end_time = datetime.strptime(end_time, "%I:%M%p").time() if ':' in end_time else datetime.strptime(end_time, "%I%p").time()
        start_datetime = datetime.combine(datetime.strptime(start_of_quarter, "%m/%d/%y").date(), start_time)
        end_datetime = datetime.combine(datetime.strptime(start_of_quarter, "%m/%d/%y").date(), end_time)

        # Map days of the week to their two-letter abbreviations
        days_of_week = [day_map[day] for day in re.findall(r'(Th|Su|M|T|W|F|S)', days_of_week)]
        
        # Remove duplicates
        days_of_week = list(dict.fromkeys(days_of_week))
        
        # Add class identifier
        title = course
        
        organizer = vCalAddress('MAILTO:wes.lauer@gmail.com')
        organizer.params['cn'] = vText(instructor)
        organizer.params['role'] = vText('Instructor')

        
        # Create the event
        event = create_event(title, start_datetime, end_datetime, days_of_week, meeting_space, organizer, classtitle)
        cal.add_component(event)

# Save calendar to file.
print(cal.to_ical().decode("utf-8"))
#print(cal)
with open('my_schedule.ics', 'wb') as my_file:
    my_file.write(cal.to_ical())
    