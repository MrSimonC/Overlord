# Overlord

## Background
Contact the BI database, run some data quality reports and email a recipient/Team.

## Prerequisite
Install "Visual C++ Redistributable for Visual Studio 2015 x86.exe" (on 32-bit, or x64 on 64-bit) which allows Python 3.5 dlls to work, found here: https://www.microsoft.com/en-gb/download/details.aspx?id=48145

## Installation and Running:
Make a folder in "C:\Program Files\Overlord" and put all the following files in there:
- overlord - windowed.exe
- Overlord Schedule.xml
- Overlord.exe
- Overlord.txt

Each file:
- overlord - windowed.exe  = not needed by the process, but is the same as Overlord.exe - but outputs to a command line, so run in cmd.exe if you have an error and want to see what's happening
- Overlord Schedule.xml = task scheduler import file
- Overlord.exe = main program
- Overlord.txt = contains the path to a folder (currently "I:\Information Systems\BACK OFFICE\Data Quality\SQL for Overlord report") which should contain:
    - Any .sql files you want to be ran
    - to.txt = lists any email addresses you want the report to be emailed to

In Windows Task Scheduler, Import Task, choose the .xml file, then change the "Run only when user is logged on" username to your own.
OR create a new task with the following attributes:
- General: Run only when user is logged on
- Trigger: at 11am everyday
- Actions: Start a program: "C:\Program Files\Overlord\overlord.exe"
- Settings:
    - Allow ask to be run on demand
    - Stop the task if it runs longer than: 4 hours
    - If the running task does not end when requested, force it stop
    - If the task is already running, then the following rule applies: Stop the existing instance

## Notes
### Outlook
Ensure your Outlook inbox doesn't fill up - else this will create error messages on your desktop when the program runs.

Written by:  
Simon Crouch, late 2016 in Python 3.5
