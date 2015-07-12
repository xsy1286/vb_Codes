Attribute VB_Name = "allStateVars"
'This is a e.g. for every state of object in this program been here
#Const App = "shot"
#If App = "" Then

#ElseIf App = "字幕ban" Then
Public top1 As Boolean
Public d As Integer
Public movable As Boolean
Public frs As Long 'Form1 get mousemouse firstly '要其它Form使用必须Public
Public p As Long
Public frm1alltop As Integer
Public loadOnce As Integer

#ElseIf App = "shot" Then
Public clip As Integer
Public dfaddress As Integer
Public iftray As Integer


#End If
