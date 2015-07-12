Attribute VB_Name = "Module1"
Public h, v As Double
Public Const a = 10
Public s As String
Public da1 As Date
Public da2 As Date
Public Function run(year As Integer) As Boolean
If year Mod 4 = 0 Then
run = True
If year Mod 100 = 0 And year Mod 400 <> 0 Then run = False
Else
run = False
End If
End Function
