Attribute VB_Name = "delay"
Option Explicit
Public Sub delayus(delaytime As Long) '看电脑速度
Dim i
For i = 1 To delaytime
DoEvents
Next i
End Sub

