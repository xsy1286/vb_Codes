Attribute VB_Name = "transparent"
Option Explicit
Public Declare Function GetWindowLong Lib "user32" Alias "GetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long) As Long
Public Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
Public Declare Function SetLayeredWindowAttributes Lib "user32" (ByVal hwnd As Long, ByVal crKey As Long, ByVal bAlpha As Byte, ByVal dwFlags As Long) As Long
Const WS_EX_LAYERED = &H80000
Const GWL_EXSTYLE = (-20)
Const LWA_COLORKEY = &H1
Public Sub transpa(TheForm As Form)
Dim transcolor As Long
 transcolor = RGB(66, 66, 66) '必须都66
 TheForm.BackColor = transcolor '必须改变自己背景
  Dim rtn As Long
rtn = GetWindowLong(TheForm.hwnd, GWL_EXSTYLE)
 rtn = rtn Or WS_EX_LAYERED
 SetWindowLong TheForm.hwnd, GWL_EXSTYLE, rtn
  SetLayeredWindowAttributes TheForm.hwnd, transcolor, 0, LWA_COLORKEY
End Sub
'ex:transpa Me
