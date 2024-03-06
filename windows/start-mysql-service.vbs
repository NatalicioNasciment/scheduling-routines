Set objShell = CreateObject("WScript.Shell")

Do
    WScript.Sleep 60000

    If Not IsServiceRunning("MySQL_PDVdb") Then

        StartService "MySQL_PDVdb"
    End If
Loop

Function IsServiceRunning(serviceName)

    Set objWMIService = GetObject("winmgmts:\.\root\cimv2")
    Set colItems = objWMIService.ExecQuery("Select * from Win32_Service where Name='" & serviceName & "'")

    If colItems.Count > 0 Then
        IsServiceRunning = True
    Else
        IsServiceRunning = False
    End If
End Function

Sub StartService(serviceName)
    objShell.Run "net start " & serviceName, 0, True
End Sub