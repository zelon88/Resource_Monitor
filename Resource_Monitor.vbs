'File Name: Resource_Monitor.vbs
'Version: v1.0, 11/14/2018
'Author: Justin Grimes, 9/27/2018

'--------------------------------------------------
'Supported Switches
' -t  -Use when run as a scheduled task. Disables output.
' -i  -Use to display as much information as possible to the user. Disables email notification.
' -ss  -Use to select which CPU socket to target for monitoring.
' -ds  -Use to select which disk to target for monitoring.
'--------------------------------------------------

'--------------------------------------------------
'Global variable definitions used in this script.
Option Explicit
Dim counter, strComputer, logDir, cacheDir, today, timeStamp, maxAge, triggerEmail, logFileName, _
 objFSO, arg, objNet, objShell, strComputerName, strUserName, mailFile, objFolder, objWMIService, _
 CPUInfo, CPUUsage, Item, loopCounter1, stringData, i, objItem, sendmail, CPUAddressWidth, _
 CPUArchitecture, CPUCaption, CPUStatus, CPUDeviceID, CPUFamily, CPUManufacturer, CPUMaxClockSpeed, _
 CPUName, CPUCoreCount, CPUProcessorCount, CPUCounter, echo1, safeDate, safeTime, dateTime, _
 memInfo1, memItem1, GB, MB, emailData, loopCounter2, AverageCPUUsage, MemoryUsage, memInfo2, _
 memItem2, TotalPhysicalMemory, TotalVirtualMemory, TotalMemory, CommittedMemory, mFile, _
 AvailablePhysicalMemory, PhysicalMemoryUtilization, emailText, logText, objLogFile, cpuInfoData, _
 cpuUsageData, memInfoData, windowTitle, serverDir, scriptsDir, objMemuCacheFile, cpuiCacheData, _
 cpuuCacheData, memuCacheData, cpuInfoCacheFileName,  memUteCacheFileName, cpuUteCacheFileName, _
 cacheCreated, logCreated, objCpuiCacheFile, objCpuuCacheFile, cacheDir1, CPUCurrentClockSpeed, _
 CPUExternalClockSpeed, loopCounter0, coreCount, CPULoadPercentage, objRefresher, colDiskDrives, _
 objDiskDrive, loopCounter3, SocketSelector, DriveName, DriveTransfersPerSecond, _
 DriveTotalTransferRate, DriveCurrentQueueLength, DriveReadsPerSecond, DriveReadBytesPerSecond, _
 DriveWritesPerSecond, DriveWriteBytesPerSecond, DriveFreeMegabytes, DriveFreeSpace, DriveBusyTime, _
 DriveIdleTime, diskInfoData, loopCounter4, DiskSelector, DriveCounter, CPUSocketCounter, diskData, _
 DriveCounter1, coreCount1, CPUUsage1, loopCounter5, cpuUsageDataCache, Problem, companyAbbreviation, _
 toEmail, fromEmail, companyName
'--------------------------------------------------

'--------------------------------------------------
'Define variables for the session
counter = maxAge = 0
companyAbbreviation = "Company"
toEmail = "IT@company.com"
fromEmail = "Server@company.com"
companyName = "Company Inc."
windowTitle = companyAbbreviation & " Resource Monitor"
strComputer = "."
serverDir = "\\Server\"
logDir = serverDir & "Logs\"
scriptsDir = serverDir & "AutomationScripts\"
cacheDir1 = scriptsDir & "Resource_Monitor\Cache\"
today = Date
timeStamp = Now
triggerEmail = FALSE
Set objFSO = CreateObject("Scripting.FileSystemObject")
Set arg = WScript.Arguments
Set objNet = CreateObject("Wscript.Network") 
Set objShell = WScript.CreateObject("WScript.Shell")
Set objWMIService = GetObject("winmgmts:\\" & strComputer & "\root\CIMV2")
strComputerName = objShell.ExpandEnvironmentStrings("%COMPUTERNAME%")
cacheDir = cacheDir1 & strComputerName & "\"
strUserName = objShell.ExpandEnvironmentStrings("%USERNAME%")
sendmail = scriptsDir & "Resource_Monitor\sendmail.exe "
mailFile = cacheDir & "Warning.mail"
safeDate = DatePart("yyyy",Date) & Right("0" & DatePart("m",Date), 2) & Right("0" & DatePart("d",Date), 2)
safeTime = Right("0" & Hour(Now), 2) & Right("0" & Minute(Now), 2) & Right("0" & Second(Now), 2)
dateTime = safeDate & "-" & safeTime
logFileName = logDir & strComputerName & "-" & dateTime & "-resource_monitor.txt"
cpuInfoCacheFileName = cacheDir & "CPUInfoCache.txt"
memUteCacheFileName = cacheDir & "MemUteCache.txt"
cpuUteCacheFileName = cacheDir & "CPUUteCache.txt"
'--------------------------------------------------

'--------------------------------------------------
'A function to create required directories if needed.
Function CreateDirectories()
  CreateDirectories = TRUE
  If Not objFSO.FolderExists(logDir) Then
    Set objFolder = objFSO.CreateFolder(logdir)
  End If
  If Not objFSO.FolderExists(cacheDir1) Then
    Set objFolder = objFSO.CreateFolder(cacheDir1)
  End If
  If Not objFSO.FolderExists(cacheDir) Then
    Set objFolder = objFSO.CreateFolder(cacheDir)
  End If
  If Not objFSO.FolderExists(cacheDir) Or Not objFSO.FolderExists(cacheDir1) Or Not objFSO.FolderExists(logDir) Then
    CreateDirectories = FALSE
  End If
End Function
'--------------------------------------------------

'--------------------------------------------------
'A function to retrieve the specified arguments.
Function GetArgs()
  GetArgs = ""
  If (arg.Count > 0) Then
    GetArgs = arg(0)
  End If
End Function
'--------------------------------------------------

'--------------------------------------------------
'A function to create an email file.
Function CreateEmail(emailData)
  If objFSO.FileExists(mailFile) Then
    objFSO.DeleteFile(mailFile)
  End If
  CreateEmail = TRUE
  Set mFile = objFSO.CreateTextFile(mailFile, TRUE, FALSE)  
  mFile.Write "To: " & toEmail & vbNewLine & "From: " & fromEmail & vbNewLine & "Subject: " & companyAbbreviation & " Low Resource Warning!!! " & _
  vbNewLine & emailData & vbNewLine & vbNewLine & "This notification was generated by " & strComputerName & "." & _
  vbNewLine & vbNewLine & "Script: ""Resource_Monitor.vbs""" 
  mFile.Close
  If Not objFSO.FileExists(sendmail) Or Not objFSO.FileExists(mailFile) Then
    CreateEmail = FALSE
  End If
End Function
'--------------------------------------------------

'--------------------------------------------------
'A function for running SendMail.
Function SendEmail() 
  SendEmail = TRUE
  If Not objFSO.FileExists(sendmail) Or Not objFSO.FileExists(mailFile) Then
    SendEmail = FALSE
  End If
  If SendEmail = TRUE Then
    objShell.run sendmail & mailFile
  End If
End Function
'--------------------------------------------------

'--------------------------------------------------
'A function to create a log file.
Function CreateLog(logData)
  logCreated = TRUE
  If Not (logData = "") Then
    Set objLogFile = objFSO.CreateTextFile(logFileName, TRUE, FALSE)
    objLogFile.WriteLine(logData)
    objLogFile.Close
  End If
  If Not objFSO.FileExists(logFileName) Then
    logCreated = FALSE
  End If
  CreateLog = logCreated
End Function
'--------------------------------------------------

'--------------------------------------------------
'A function to get disk information.
Function GetDiskInfo(Stopper)
  Set objRefresher = CreateObject("WbemScripting.SWbemRefresher")
  Set colDiskDrives = objRefresher.AddEnum(objWMIService, "Win32_PerfFormattedData_PerfDisk_LogicalDisk").objectSet
  loopCounter4 = 0
  objRefresher.Refresh
  For Each objDiskDrive in colDiskDrives
    If Stopper = 0 Then
      Exit For
    End If
    DriveName = objDiskDrive.Name
    DriveTransfersPerSecond = objDiskDrive.DiskTransfersPerSec
    DriveTotalTransferRate = objDiskDrive.DiskBytesPerSec
    DriveCurrentQueueLength = objDiskDrive.CurrentDiskQueueLength
    DriveReadsPerSecond = objDiskDrive.DiskReadsPerSec
    DriveReadBytesPerSecond = objDiskDrive.DiskReadBytesPerSec
    DriveWritesPerSecond = objDiskDrive.DiskWritesPerSec
    DriveWriteBytesPerSecond = objDiskDrive.DiskWriteBytesPerSec
    DriveFreeMegabytes = objDiskDrive.FreeMegabytes
    DriveFreeSpace = objDiskDrive.PercentFreeSpace
    DriveBusyTime = objDiskDrive.PercentDiskTime
    DriveIdleTime = objDiskDrive.PercentIdleTime
    Stopper = Stopper - 1
    loopCounter4 = loopCounter4 + 1
    DriveCounter = loopCounter4
  Next
End Function 
'--------------------------------------------------

'--------------------------------------------------
'A function to display disk information.
Function DisplayDiskInfo()
  diskInfoData = strComputerName & vbNewLine & "Disk Specifications" & vbNewLine & vbNewLine & _
   "Drive Label: " & DriveName & vbNewLine & _
   "Total Transfers Per Second: " & DriveTransfersPerSecond & vbNewLine & _
   "Total Bytes Per Second: " & DriveTotalTransferRate & vbNewLine & _
   "Current Queue Length: " & DriveCurrentQueueLength & vbNewLine & _
   "Reads Per Second: " & DriveReadsPerSecond & vbNewLine & _
   "Read Bytes Per Second: " & DriveReadBytesPerSecond & vbNewLine & _
   "Writes Per Second: " & DriveWritesPerSecond & vbNewLine & _
   "Write Bytes Per Second: " & DriveWriteBytesPerSecond & vbNewLine & _
   "Free Space (MB): " & DriveFreeMegabytes & vbNewLine & _
   "Free Space (%): " & DriveFreeSpace & "%" & vbNewLine & _
   "Busy Time: " & DriveBusyTime & vbNewLine & _
   "Idle Time: " & DriveIdleTime & vbNewline & _
   "Disk Number: " & DriveCounter & vbNewline
  DisplayDiskInfo = diskInfoData
  MsgBox diskInfoData, 64, windowTitle
End Function
'--------------------------------------------------

'--------------------------------------------------
'A function to retrieve processor information.
'Set "Stopper" to the CPU socket that you want information about. 
  'If "Stopper" is -1 or not set, the highest socket will be used.
  'Socket numbering starts at 0. So valid input for 2 sockets is 0 or 1.
Function GetProcessorInfo(Stopper)
  Const wbemFlagReturnImmediately = &h10
  Const wbemFlagForwardOnly = &h20
  loopCounter3 = 0
  Set CPUInfo = objWMIService.ExecQuery("SELECT * FROM Win32_Processor", "WQL", wbemFlagReturnImmediately + wbemFlagForwardOnly)
  For Each objItem In CPUInfo
    If Stopper = 0 Then
      Exit For
    End If
    CPUAddressWidth = objItem.AddressWidth
    CPUArchitecture = objItem.Architecture
    CPUCaption = objItem.Caption
    CPUStatus = objItem.CpuStatus
    CPUDeviceID = objItem.DeviceID
    CPUFamily = objItem.Family
    CPUManufacturer = objItem.Manufacturer
    CPUCurrentClockSpeed = objItem.CurrentClockSpeed
    CPUMaxClockSpeed = objItem.MaxClockSpeed
    CPUExternalClockSpeed = objItem.ExtClock
    CPUName = objItem.Name
    CPULoadPercentage = objItem.LoadPercentage
    CPUCoreCount = objItem.NumberOfCores
    CPUProcessorCount = objItem.NumberOfLogicalProcessors
    Stopper = Stopper - 1
    loopCounter3 = loopCounter3 + 1
    CPUSocketCounter = loopCounter3
  Next
  If CPUArchitecture = 0 Then
    CPUArchitecture = "x86"
  End If
  If CPUArchitecture = 1 Then
    CPUArchitecture = "MIPS"
  End If
  If CPUArchitecture = 2 Then
    CPUArchitecture = "Alpha"
  End If
  If CPUArchitecture = 3 Then
    CPUArchitecture = "PowerPC"
  End If
  If CPUArchitecture = 5 Then
    CPUArchitecture = "ARM"
  End If
  If CPUArchitecture = 6 Then
    CPUArchitecture = "ia64"
  End If
  If CPUArchitecture = 9 Then
    CPUArchitecture = "x64"
  End If
End Function
'--------------------------------------------------

'--------------------------------------------------
'A function to display processor information.
Function DisplayProcessorInfo()
  cpuInfoData = strComputerName & vbNewLine & "CPU Specifications" & vbNewLine & vbNewLine & _
   "CPU Address Width: " & CPUAddressWidth & vbNewLine & _
   "CPU Architecture: " & CPUArchitecture & vbNewLine & _
   "CPU Description: " & CPUCaption & vbNewLine & _
   "CPU Status: " & CPUStatus & vbNewLine & _
   "CPU Device ID: " & CPUDeviceID & vbNewLine & _
   "CPU Family: " & CPUFamily & vbNewLine & _
   "CPU Manufacturer: " & CPUManufacturer & vbNewLine & _
   "CPU Current Clockspeed: " & CPUCurrentClockSpeed & vbNewLine & _
   "CPU Max Clockspeed: " & CPUMaxClockSpeed & vbNewLine & _
   "CPU External Clockspeed: " & CPUExternalClockSpeed & vbNewLine & _
   "CPU Name: " & CPUName & vbNewLine & _
   "CPU Load Percentage: " & CPULoadPercentage & vbNewline & _
   "CPU Physical Core Count: " & CPUCoreCount & vbNewLine & _
   "CPU Logical Thread Count: " & CPUProcessorCount & vbNewLine & _
   "CPU Socket Number: " & CPUSocketCounter & vbNewLine
  DisplayProcessorInfo = cpuInfoData
  MsgBox cpuInfoData, 64, windowTitle
End Function
'--------------------------------------------------

'--------------------------------------------------
'A function to retrieve processor usage.
Function GetProcessorUsage()
  Set CPUUsage = objWMIService.ExecQuery("SELECT * FROM Win32_PerfFormattedData_PerfOS_Processor",,48) 
  loopCounter0 = 0
  For Each Item in CPUUsage
    loopCounter0 = loopCounter0 + 1
  Next
  GetProcessorUsage = loopCounter0
End Function
'--------------------------------------------------

'--------------------------------------------------
'A function to display processor utilization to the user.
Function DisplayProcessorUsage()
  coreCount = GetProcessorUsage()
  Set CPUUsage = objWMIService.ExecQuery("SELECT * FROM Win32_PerfFormattedData_PerfOS_Processor",,48) 
  loopCounter1 = 1
  cpuUsageData = strComputerName & vbNewLine & "CPU Usage Information" & vbNewLine & vbNewLine
  For Each Item in CPUUsage
    echo1 = loopCounter1
    If coreCount = loopCounter1 Then
      echo1 = "Average"
    End If
    cpuUsageData = cpuUsageData & "Processor Usage " & echo1 & ": " & Item.PercentProcessorTime & "%" & vbNewLine
    loopCounter1 = loopCounter1 + 1
  Next
  DisplayProcessorUsage = cpuUsageData
  MsgBox cpuUsageData, 64, windowTitle
End Function
'--------------------------------------------------

'--------------------------------------------------
'A function to display processor utilization to the user.
Function DisplayProcessorUsageCache() 
  coreCount1 = GetProcessorUsage()
  Set CPUUsage1 = objWMIService.ExecQuery("SELECT * FROM Win32_PerfFormattedData_PerfOS_Processor",,48) 
  loopCounter5 = 1
  For Each Item in CPUUsage1
    cpuUsageDataCache = cpuUsageDataCache &  Item.PercentProcessorTime & vbNewLine
    loopCounter5 = loopCounter5 + 1
  Next
  DisplayProcessorUsageCache = cpuUsageDataCache
End Function
'--------------------------------------------------

'--------------------------------------------------
'A function to retrieve memory usage and information. 
Function GetMemoryInfo()
  GB = 1024 *1024 * 1024
  MB = 1024 *1024
  Set memInfo1 = objWMIService.ExecQuery("Select * from Win32_PerfFormattedData_PerfOS_Memory",,48)
  For Each memItem1 in memInfo1
    TotalPhysicalMemory = Round(Round(memItem1.AvailableBytes / GB,3) + Round(memItem1.CommittedBytes / GB,3),3)
    TotalVirtualMemory = Round(Round(memItem1.CommitLimit / GB,3) - Round(Round(memItem1.AvailableBytes / GB,3) + Round(memItem1.CommittedBytes / GB,3)),3)
    TotalMemory = Round(memItem1.CommitLimit / GB,3)
    CommittedMemory = Round(memItem1.CommittedBytes / GB,3)
    AvailablePhysicalMemory = Round(memItem1.AvailableBytes / GB,3)
    PhysicalMemoryUtilization = Round(Round(Round(memItem1.CommittedBytes / GB,3),3) / Round(Round(memItem1.AvailableBytes / GB,3) + Round(memItem1.CommittedBytes / GB,3),3) * 100,3)
  Next
End Function
'--------------------------------------------------

'--------------------------------------------------
'A function to display memory usage and information.
Function DisplayMemoryInfo()
  memInfoData = ""
  GB = 1024 *1024 * 1024
  MB = 1024 *1024
  Set memInfo2 = objWMIService.ExecQuery("Select * from Win32_PerfFormattedData_PerfOS_Memory",,48)
  For Each memItem2 in memInfo2
    memInfoData = strComputerName & vbNewLine & "Memory Usage Information" & vbNewLine & vbNewLine & _
     "Total Physical Memory: " & TotalPhysicalMemory & " GB" & vbNewLine & _
     "Total Virtual Memory: " & TotalVirtualMemory & " GB" & vbNewLine & _
     "Total Memory: " & TotalMemory & " GB" & vbNewLine & _
     "Committed Memory: " & CommittedMemory & " GB" & vbNewLine & _
     "Available Physical Memory: " & AvailablePhysicalMemory & " GB" & vbNewLine & _
     "Physical Memory Usage: " & PhysicalMemoryUtilization & "%" & vbNewLine
    DisplayMemoryInfo = memInfoData
    MsgBox memInfoData, 64, windowTitle
  Next
End Function
'--------------------------------------------------

'--------------------------------------------------
'A function to detect if the machine's disks are busy.
Function DetectDiskBusy()
  diskData = ""
  Problem = FALSE
  DiskSelector = -1
  If GetArgs() = "-ds" Then
    DiskSelector = InputBox("Disk Selector" & vbNewLine & vbNewline & "Please select which disk you would like to monitor.",windowTitle,,64)
  End If
  GetDiskInfo(DiskSelector)
  If DriveIdleTime <= 25 Then
    Problem = "Drive Idle Time: " & DriveIdleTime
  End If
  If DriveBusyTime >= 75 Then
    Problem = "Drive Busy Time: " & DriveBusyTime
  End If
  If Replace(DriveFreeSpace,"%","") <= 15 Then
    Problem = "Percent Free Disk Space: " & DriveFreeSpace
  End If
  DetectDiskBusy = Problem
End Function
'--------------------------------------------------

'--------------------------------------------------
'A function to detect if the machine's CPU is busy.
Function DetectCPUBusy()
  SocketSelector = -1
  If GetArgs() = "-ss" Then
    SocketSelector = InputBox("Socket Selector" & vbNewLine & vbNewline & "Please select which CPU socket you would like to monitor.",windowTitle,,64)
  End If
  GetProcessorInfo(SocketSelector)
  coreCount = GetProcessorUsage()
  Set CPUUsage = objWMIService.ExecQuery("SELECT * FROM Win32_PerfFormattedData_PerfOS_Processor",,48) 
  loopCounter2 = 0
  DetectCPUBusy = FALSE
  For Each Item in CPUUsage
    If coreCount = loopCounter2 Then
      AverageCPUUsage = Item.PercentProcessorTime
    End If
    loopCounter2 = loopCounter2 + 1
  Next
  DetectCPUBusy = AverageCPUUsage
End Function
'--------------------------------------------------

'--------------------------------------------------
'A function to detect if the machine's memory is busy.
Function DetectMemoryBusy()
  DetectMemoryBusy = FALSE
  GetMemoryInfo()
  DetectMemoryBusy = PhysicalMemoryUtilization
End Function
'--------------------------------------------------

'--------------------------------------------------
'A function to save all the data gathered into cache files that can be retrieved by other applications.
Function CreateCache()
  cacheCreated = TRUE
  If Not objFSO.FolderExists(cacheDir) Then
    cacheCreated = FALSE
  End If
  If objFSO.FileExists(memUteCacheFileName) Then
    objFSO.DeleteFile(memUteCacheFileName)
  End If
  memuCacheData = TotalPhysicalMemory & vbNewLine & TotalVirtualMemory & vbNewLine & TotalMemory & vbNewLine & _
   CommittedMemory & vbNewline & AvailablePhysicalMemory & vbNewline & PhysicalMemoryUtilization
  Set objMemuCacheFile = objFSO.CreateTextFile(memUteCacheFileName, TRUE, FALSE)
  objMemuCacheFile.WriteLine(memuCacheData)
  objMemuCacheFile.Close
  If objFSO.FileExists(cpuUteCacheFileName) Then
    objFSO.DeleteFile(cpuUteCacheFileName)
  End If
  cpuuCacheData = DisplayProcessorUsageCache
  Set objCpuuCacheFile = objFSO.CreateTextFile(cpuUteCacheFileName, TRUE, FALSE)
  objCpuuCacheFile.WriteLine(cpuuCacheData)
  objCpuuCacheFile.Close
  If objFSO.FileExists(cpuInfoCacheFileName) Then
    objFSO.DeleteFile(cpuInfoCacheFileName)
  End If
  cpuiCacheData = CPUAddressWidth & vbNewline & CPUArchitecture & vbNewline & CPUCaption & vbNewline & _
   CPUStatus & vbNewline & CPUDeviceID & vbNewline & CPUFamily & vbNewline & CPUManufacturer & vbNewline & _
   CPUCurrentClockSpeed & vbNewline & CPUMaxClockSpeed & vbNewline & CPUExternalClockSpeed & vbNewline & _
   CPUName & vbNewline & CPULoadPercentage & vbNewLine & CPUCoreCount & vbNewline & CPUProcessorCount
  Set objCpuiCacheFile = objFSO.CreateTextFile(cpuInfoCacheFileName, TRUE, FALSE)
  objCpuiCacheFile.WriteLine(cpuiCacheData)
  objCpuiCacheFile.Close
  If Not objFSO.FileExists(memUteCacheFileName) Or Not objFSO.FileExists(cpuUteCacheFileName) Or Not objFSO.FileExists(cpuInfoCacheFileName) Then
    cacheCreated = FALSE
  End If
  CreateCache = cacheCreated
End Function
'--------------------------------------------------

'--------------------------------------------------
'The main function of the application which makes use of the functions above.
If CreateDirectories() = TRUE Then
  'When run as a task: detect if the machine's memory is busy and notify IT.
  If DetectMemoryBusy() >= 75 Then
    logText = strComputerName & " is running low on memory resources." & VBNewLine & VBNewLine & "Memory Usage: " & DetectMemoryBusy & "%"
    'Attempt to create a log file and return an error upon failure.
    If CreateLog(logText) = FALSE Then
      MsgBox "ERROR!!! RSM308, Cannot create a log file on " & timeStamp & "!", 16, windowTitle
    End If
    'Check if the -t argument is set for running as a task.
    If GetArgs() = "-t" Then
      emailText = "This is an automated message from the " & companyName & " Network to notify you that " & strComputerName & " is running low on memory resources." & _
       VBNewLine & VBNewLine & "Memory Usage: " & DetectMemoryBusy & "%"
      'Attempt to create an email file and return an error upon failure.
      If CreateEmail(emailText) = FALSE Then
        MsgBox "ERROR!!! RSM316, Cannot create an email file on " & timeStamp & "!", 16, windowTitle
      End If
      'Attempt to send a notification email and return an error upon failure.
      If SendEmail() = FALSE Then
        MsgBox "ERROR!!! RSM321, Cannot send a notification email on " & timeStamp & "!", 16, windowTitle
      End If
    End If
  End If
  'When run as a task: detect if the machine's CPU is busy and notify IT.
  If DetectCPUBusy() >= 75 Then
    logText = strComputerName & " is running low on memory resources." & VBNewLine & VBNewLine & "CPU Usage: " & DetectCPUBusy & "%"
    If CreateLog(logText) = FALSE Then
      MsgBox "ERROR!!! RSM328, Cannot create a log file on " & timeStamp & "!", 16, windowTitle
    End If
    If GetArgs() = "-t" Then
      emailText = "This is an automated message from the " & companyName & " Network to notify you that " & strComputerName & " is running low on CPU resources." & _
       VBNewLine & VBNewLine & "CPU Usage: " & DetectCPUBusy & "%"
      'Attempt to create an email file and return an error upon failure.
      If CreateEmail(emailText) = FALSE Then
        MsgBox "ERROR!!! RSM335, Cannot create an email file on " & timeStamp & "!", 16, windowTitle
      End If
      'Attempt to send a notification email and return an error upon failure.
      If SendEmail() = FALSE Then
        MsgBox "ERROR!!! RSM338, Cannot send a notification email on " & timeStamp & "!", 16, windowTitle
      End If
    End If
  End If
  'When run as a task: detect if the machine's disks are busy and notify IT.
  If DetectDiskBusy() <> FALSE And DetectDiskBusy() <> NULL Then
    logText = strComputerName & " is running with stressed disk resources." & VBNewLine & VBNewLine & DetectDiskBusy 
    If CreateLog(logText) = FALSE Then
      MsgBox "ERROR!!! RSM471, Cannot create a log file on " & timeStamp & "!", 16, windowTitle
    End If
    If GetArgs() = "-t" Then
      emailText = "This is an automated message from the " & companyName & " Network to notify you that " & strComputerName & " is running with stressed disk resources." & _
       VBNewLine & VBNewLine &  DetectDiskBusy 
      'Attempt to create an email file and return an error upon failure.
      If CreateEmail(emailText) = FALSE Then
        MsgBox "ERROR!!! RSM478, Cannot create an email file on " & timeStamp & "!", 16, windowTitle
      End If
      'Attempt to send a notification email and return an error upon failure.
      If SendEmail() = FALSE Then
        MsgBox "ERROR!!! RSM482, Cannot send a notification email on " & timeStamp & "!", 16, windowTitle
      End If
    End If
  End If
  'When run by a user: display as much information to the user as possible.
  If GetArgs() <> "-t" Then
    DisplayProcessorInfo()
    DisplayProcessorUsage()
    DisplayMemoryInfo()
    DisplayDiskInfo()
  End If
  If CreateCache() = FALSE Then
    MsgBox "ERROR!!! RSM347, Cannot create a cache file on " & timeStamp & "!", 16, windowTitle
  End If
End If
'--------------------------------------------------
