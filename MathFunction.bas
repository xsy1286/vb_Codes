Attribute VB_Name = "MathFunction"
Option Explicit
Public Function order(ByRef a As Long, ByRef b As Long) As Long
    
    Dim c As Long
    If (a > b) Then
        c = a
        a = b
        b = c
    End If

order = 1

End Function
