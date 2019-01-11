# Resource_Monitor
A Windows Resource Monitor in VBS with email alerts. Can run on-demand, from command line, or as a task.

## Attributes
NAME: Resource_Monitor.vbs

TYPE: Visual Basic Script

PRIMARY LANGUAGE: VBS

AUTHOR: Justin Grimes ([@zelon88](https://github.com/zelon88))

ORIGINAL VERSION DATE: 7/16/2018

CURRENT VERSION DATE: 11/14/2018

VERSION: v1.0


## DESCRIPTION: 
A simple script for monitoring the resource usage of workstations or servers and notifying IT if and when there is a problem.


## PURPOSE: 
To ensure smooth operation of the network and prevent downtime by creating visibility on device health for administrators.


## INSTALLATION INSTRUCTIONS: 
1. Copy the entire "Resource_Monitor" folder into the "AutomationScripts" folder on SERVER (or any other network accessible location).
2. Add a scheduled task to SERVER to run "Resource_Monitor.vbs" several times daily with the "-t" argument.
3. Use the -t argument when running from Task Scheduler. This will send emails to IT, warning them about resource problems.
4. Do not use the -t argument when running as a user. This will prevent emails from being sent and display the information instead. DOES NOT DISABLE LOGGING!
5. -i  -Use to display as much information as possible to the user. Disables email notification.
6. -ss  -Use to select which CPU socket to target for monitoring.
7. -ds  -Use to select which disk to target for monitoring.

NOTES: SendMail for Windows is required and included in the "Resource_Monitor" folder. The SendMail data files must be included in the same directory as "Resource_Monitor.vbs" in order for emails to be sent correctly.
