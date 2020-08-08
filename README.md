## Pre requisites:
-   pip install win32com.client
-   pip install datetime
-   pip install os
-   pip install subprocess
-   pip install time
-   pip install xlrd

Python script to automate meeting schedules in outlook, wherein it sends mail to the recipients configured in excel automatically at the set time
The Script currently works only if run manually. It can be run at set time by creating a .bat 🦇 file and creating a new Task scheduler process.

## Creation of .bat file and Task Scheduler
- Create a .bat file in the current workind directory and edit it by giving the path of python executable and path to the seatBook.py 
    Should look something like this :
    #
    "C:\Users\User\AppData\Local\Programs\Python\Python38-32\python.exe" "C:\Users\User\ 'path to seatBook.py'"
    exit 0
- Open task scheduler and create a task. The location of the .bat file should be given under "actions" tab. Give the desired time under "Triggers"
