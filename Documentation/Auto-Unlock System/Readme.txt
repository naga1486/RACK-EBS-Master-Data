Auto-unlock system:
==================
In many projects, we need to schedule the execution of automation scripts using Windows scheduler. However, this would not work if the system is locked. This problem can be solved with the auto-unlock utility called Logon.exe. This utility will automatically unlock the system at the scheduled time, so that the QTP scripts can be kicked off without any issues. One limitation of this, however, is that it may not work for machines accessed using Citrix/Remote Desktop. 

Refer the file "Logon command line utility.htm" for more details. This utility can be downloaded from: http://www.softtreetech.com/24x7/archive/51.htm.

Steps to use:
------------
* Open Logon.bat in notepad and change the path appropriately (lines 1 and 2), to point to "Logon.exe" folder
* Enter your password in line 3, which will be used to unlock the system at the scheduled time
* Update the path of the "InitScript_RACK.vbs" as approriate (lines 4 and 5)
* Open Windows Scheduler and schedule Logon.bat as required
* Bingo :)
