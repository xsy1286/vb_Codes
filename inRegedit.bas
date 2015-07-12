Attribute VB_Name = "inRegedit"
Private WshShell As Object
Private exetemp As String
Public Function setStartUp(b As Boolean, AppFilename As String) As Long
On Error GoTo Errlog
    Set WshShell = CreateObject("wscript.shell")
    exetemp = IIf(Right(App.Path, 1) = "\", App.Path, App.Path & "\") & App.EXEName & ".exe"
    WshShell.regwrite "HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\Windows\CurrentVersion\Run\", exetemp '加入到注册表，开机运行
 
Exit Function

Errlog:
 Call whenErr(Err.Number, AppFilename, "setStartUp(Bool)")

End Function
