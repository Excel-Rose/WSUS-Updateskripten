# WSUS-Updateskripten

Script to remotely run updates: as the WU Agent cannot be run remotely, a task is disctributed via GPO and triggered remotely.
The script starts a clients using WOL using a csv-file, installs the updates, reboots the client and triggers a report.

I rewrote the whole script using workflows to speed up the process, so instead of updating all the clients one after the other,
4 are updates simultaneously.

The script was tested on a Server 2016 machine with Windows 10 Enterprise clients.

This is my first "bigger" Powershell project! Any feedback is highly welcome! If you copy and run parts of the script,
you do so at your own risk!
