Attribute VB_Name = "MoveForm"
Public Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long

Public Declare Function ReleaseCapture Lib "user32" () As Long

Public Sub DoDrag(TheForm As Form)
    If TheForm.WindowState <> vbMaximized Then
        ReleaseCapture
        SendMessage TheForm.hwnd, &HA1, 2, 0&
    End If
End Sub
'ex: Private Sub Text1_MouseDown(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
'     DoDrag Me
'    End Sub
