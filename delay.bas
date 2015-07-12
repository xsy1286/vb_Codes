Attribute VB_Name = "delay"
Option Explicit
Public Sub delayus(delaytime As Long) '看电脑速度
Dim i As Long
For i = 1 To delaytime
DoEvents
Next i
End Sub

Public Function waittime(delay As Single)
  Dim starttime As Single
  starttime = Timer
  Do Until (Timer - starttime) > delay
  DoEvents
  Loop
End Function
