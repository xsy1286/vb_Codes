Attribute VB_Name = "Object_location"
Public Sub mid_obj(theobj As Object, TheForm As Form)
theobj.Left = (TheForm.Width - theobj.Width) / 2
theobj.Top = (TheForm.Height - theobj.Height) / 2
End Sub

Public Sub angle_obj(theobj As Object, TheForm As Form, num As Integer)
Select Case (num)
Case 1:
theobj.Left = TheForm.Width - theobj.Width
theobj.Top = 0
Case 2:
theobj.Left = TheForm.Width - theobj.Width
theobj.Top = TheForm.Height - theobj.Height
Case 3:
theobj.Left = 0
theobj.Top = TheForm.Height - theobj.Height
Case 4:
theobj.Left = 0
theobj.Top = 0
End Select

End Sub


Public Sub tophalf_obj(theobj As Object, TheForm As Form)
theobj.Left = (TheForm.Width - theobj.Width) / 2
theobj.Top = TheForm.Height / 2 - theobj.Height
End Sub
Public Sub endhalf_obj(theobj As Object, TheForm As Form)
theobj.Left = (TheForm.Width - theobj.Width) / 2
theobj.Top = TheForm.Height / 2
End Sub
