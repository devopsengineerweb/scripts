Option Explicit
Call KillProcessbyName("common-api.jar")
'*********************************************************************************
Sub KillProcessbyName(FileName)
    On Error Resume Next
    Dim WshShell,strComputer,objWMIService,colProcesses,objProcess
    Set WshShell = CreateObject("Wscript.Shell")
    strComputer = "."
    Set objWMIService = GetObject("winmgmts:{impersonationLevel=impersonate}!\\" & strComputer & "\root\cimv2")
    Set colProcesses = objWMIService.ExecQuery("SELECT * FROM Win32_Process")
    For Each objProcess in colProcesses
        If InStr(objProcess.CommandLine,FileName) > 0 Then
            If Err <> 0 Then
                MsgBox Err.Description,VbCritical,Err.Description
            Else
                objProcess.Terminate(0) 
            End if
        End If
    Next
End Sub
'**********************************************************************************
