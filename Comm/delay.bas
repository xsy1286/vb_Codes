Attribute VB_Name = "delay"
Option Explicit
Public Sub delayus(delaytime As Long) '�������ٶ�
Dim i
For i = 1 To delaytime
DoEvents
Next i
End Sub

