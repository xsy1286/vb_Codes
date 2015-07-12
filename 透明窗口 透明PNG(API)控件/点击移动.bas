Attribute VB_Name = "Module2"
Public Declare Function SetWindowPos Lib "user32" (ByVal hwnd As Long, ByVal hWndInsertAfter As Long, ByVal x As Long, ByVal y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long) As Long
Public Declare Function GetCursorPos Lib "user32" (lpPoint As POINTAPI) As Long
Public Type POINTAPI
x As Long
y As Long
End Type
Public p As POINTAPI
Public p2 As POINTAPI
Public t As Integer



'要点击移动控件  Mouse代码
'Private Sub Form_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
't = 1
'p2.x = x '/ Screen.TwipsPerPixelX
'p2.y = y '/ Screen.TwipsPerPixelY
'Timer2.Interval = 20
'Timer2.Enabled = True
'End Sub

'Private Sub Form_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
't = 0
'Timer2.Enabled = False
'End Sub


'添加timer2控件
'timer代码
'Private Sub Timer2_Timer()
'If t = 1 Then GetCursorPos p
'Me.Left = (p.x - p2.x) * Screen.TwipsPerPixelX
'Me.Top = (p.y - p2.y) * Screen.TwipsPerPixelY
'End Sub
