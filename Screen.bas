Attribute VB_Name = "screen_attribute"
Public Declare Function GetSystemMetrics Lib "user32" (ByVal nIndex As Long) As Long
Public Const SM_CXSCREEN = 0
Public Const SM_CYSCREEN = 1
Private Declare Function SetWindowPos Lib "user32" (ByVal hwnd As Long, ByVal hWndInsertAfter As Long, ByVal X As Long, ByVal Y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long) As Long
Const HWND_TOPMOST = -1
Const HWND_NOTOPMOST = -2
Const SWP_NOSIZE = &H1
Const SWP_NOMOVE = &H2
Const SWP_NOACTIVATE = &H10
Const SWP_SHOWWINDOW = &H40
Const HWND_TOP = 0
Const SWP_NOREDRAW = &H8
Const SWP_NOREPOSITION = &H200
Private Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lparam As Any) As Long
Private Declare Function ReleaseCapture Lib "user32" () As Long

Public Sub DoDrag(TheForm As Form)
    If TheForm.WindowState <> vbMaximized Then
        ReleaseCapture
        SendMessage TheForm.hwnd, &HA1, 2, 0&
        
    End If
    
'在该控件的方法调用:
'Private Sub Label1_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
'DoDrag Me
'End Sub

End Sub



'Me.Width = GetSystemMetrics(SM_CXSCREEN) * Screen.TwipsPerPixelX 'screen.width必须用此替代
'Me.Height = GetSystemMetrics(SM_CYSCREEN) * Screen.TwipsPerPixelY

'--------------------------------------------------------------------------------------------
Public Sub mid_Form(TheForm As Form)

TheForm.Left = (GetSystemMetrics(SM_CXSCREEN) * Screen.TwipsPerPixelX - TheForm.Width) / 2
TheForm.top = (GetSystemMetrics(SM_CYSCREEN) * Screen.TwipsPerPixelY - TheForm.Height) / 2

End Sub
Public Sub top_hWnd(h As Long, Optional top As Boolean = True)
If top = True Then
SetWindowPos h, HWND_TOPMOST, 0, 0, 0, 0, SWP_NOMOVE Or SWP_NOSIZE
Else
SetWindowPos h, HWND_NOTOPMOST, 0, 0, 0, 0, SWP_NOMOVE Or SWP_NOSIZE
End If
End Sub
Public Sub top_Form(TheForm As Form, Optional top As Boolean = True)
If top = True Then
SetWindowPos TheForm.hwnd, HWND_TOPMOST, 0, 0, 0, 0, SWP_NOMOVE Or SWP_NOSIZE
Else
SetWindowPos TheForm.hwnd, HWND_NOTOPMOST, 0, 0, 0, 0, SWP_SHOWWINDOW Or SWP_NOMOVE Or SWP_NOSIZE
End If
End Sub
Public Sub all_Screen(ByVal TheForm As Form)
 TheForm.WindowState = 2
 
'TheForm.Left = 0
'TheForm.top = 0
'TheForm.Width = GetSystemMetrics(SM_CXSCREEN) * Screen.TwipsPerPixelX
'TheForm.Height = GetSystemMetrics(SM_CYSCREEN) * Screen.TwipsPerPixelY
End Sub

Public Function get_screen(ByRef X As Long, ByRef Y As Long) As Long
X = GetSystemMetrics(SM_CXSCREEN) * Screen.TwipsPerPixelX
Y = GetSystemMetrics(SM_CXSCREEN) * Screen.TwipsPerPixelY
get_screen = 1
End Function
