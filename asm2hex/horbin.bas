Attribute VB_Name = "horbin"
Public Function vtoh(ByVal a As Integer, ByVal bb As Integer) As String
Dim kkk As Integer
Dim k As Integer
Dim sixteen2 As String
sixteen2 = ""

kkk = bb

Do While (kkk > 0)

k = ((a Mod (16 ^ kkk)) - (a Mod (16 ^ (kkk - 1)))) / (16 ^ (kkk - 1))

Debug.Print "kkk= " & CStr(kkk)
Debug.Print "k= " & CStr(k)

If (k < 10) Then
  sixteen2 = CStr(k)
Else
 Select Case k
 Case 10
 sixteen2 = "A"
 Case 11
 sixteen2 = "B"
  Case 12
 sixteen2 = "C"
  Case 13
 sixteen2 = "D"
  Case 14
 sixteen2 = "E"
  Case 15
 sixteen2 = "F"
 End Select
End If

vtoh = vtoh & sixteen2
kkk = kkk - 1
Loop

End Function


