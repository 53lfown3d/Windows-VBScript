' Runs against a remote computer and outputs the WSUS registry settings and
' whether or not the machine needs a post-update reboot

strComputer = WScript.Arguments.Item(0)

Set objWMIService = GetObject("winmgmts:\\" & strComputer & "\root\cimv2")

Set objAutoUpdate = CreateObject("Microsoft.Update.AutoUpdate")
Wscript.Echo "Service enabled: " & objAutoUpdate.ServiceEnabled

Set objSettings = objAutoUpdate.Settings

Wscript.Echo "Notification level: " & objSettings.NotificationLevel
Wscript.Echo "Read-only: " & objSettings.ReadOnly
Wscript.Echo "Required: " & objSettings.Required
Select Case objSettings.ScheduledInstallationDay
    Case 0
        Wscript.Echo "Scheduled installation day: Every day"
    Case 1
        Wscript.Echo "Scheduled installation day: Sunday"
    Case 2
        Wscript.Echo "Scheduled installation day: Monday"
    Case 3
        Wscript.Echo "Scheduled installation day: Tuesday"
    Case 4
        Wscript.Echo "Scheduled installation day: Wednesday"
    Case 5
        Wscript.Echo "Scheduled installation day: Thursday"
    Case 6
        Wscript.Echo "Scheduled installation day: Friday"
    Case 7
        Wscript.Echo "Scheduled installation day: Saturday"
    Case Else
        Wscript.Echo "The scheduled installation day could not be determined."
End Select

If objSettings.ScheduledInstallationTime = 0 Then
    Wscript.Echo "Scheduled installation time: 12:00 AM"
ElseIf objSettings.ScheduledInstallationTime = 12 Then
    Wscript.Echo "Scheduled installation time: 12:00 PM"
Else
    If objSettings.ScheduledInstallationTime > 12 Then
        intScheduledTime = objSettings.ScheduledInstallationTime - 12
        strScheduledTime = intScheduledTime & ":00 PM"
    Else
        strScheduledTime = objSettings.ScheduledInstallationTime & ":00 AM"
    End If
    Wscript.Echo "Scheduled installation time: " & strScheduledTime
End If

Set objSysInfo = CreateObject("Microsoft.Update.SystemInfo")
If objSysInfo.RebootRequired Then
    Wscript.Echo "This computer needs to be rebooted."
Else
    Wscript.Echo "This computer does not need to be rebooted."
End If
