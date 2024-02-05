This is a repo for teaching utilities.

Currently, the only project is the "Calendars" project, which includes code for
converting a spreadsheet version of a course schedule into an ICS schedule
file that can be imported into Outlook. To import the file, select "File" then
"Open and Export", then import the file as a new calendar.  

If separate calendars are desired for each department or course code, separate
spreadsheets should be used.  However, for checking course conflicts, all
courses should be in the same spreadsheet.

To check course conflicts, the spreadsheet includes a location to enter the
course schedules for which conflicts are to be determined.  As many as desired
can be entered.  Then the course saves the conflicting times to a textfile
and also printes them out to a message box.

The program requires the ics library, which can be installed with PIP or through
the Anaconda.
