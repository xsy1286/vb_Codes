Attribute VB_Name = "modSystray"
'******************************
'  源码学习下载www.lvcode.com
'    欢迎分享源码给Love代码
'******************************

'Systray Module
Option Explicit

Public blnClick                  As Boolean
Public vbTray                    As NOTIFYICONDATA

Public Declare Function Shell_NotifyIcon Lib "shell32" Alias "Shell_NotifyIconA" (ByVal dwMessage As Long, pnid As NOTIFYICONDATA) As Boolean
Public Declare Function SetForegroundWindow Lib "user32" (ByVal hWnd As Long) As Long
Public Declare Function SetWindowPos Lib "user32" (ByVal hWnd As Long, ByVal hWndInsertAfter As Long, ByVal X As Long, ByVal Y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long) As Long
    
Public Type NOTIFYICONDATA
    cbSize                       As Long
    hWnd                         As Long
    uId                          As Long
    uFlags                       As Long
    uCallBackMessage             As Long
    hIcon                        As Long
    szTip                        As String * 64
End Type

Public Const SWP_NOMOVE          As Long = &H2
Public Const SWP_NOSIZE          As Long = &H1
Public Const FLAGS               As Long = SWP_NOMOVE Or SWP_NOSIZE
Public Const WM_RBUTTONUP        As Long = &H205
Public Const WM_RBUTTONCLK       As Long = &H204
Public Const WM_LBUTTONCLK       As Long = &H202
Public Const WM_LBUTTONDBLCLK    As Long = &H203
Public Const WM_MOUSEMOVE        As Long = &H200
Public Const NIM_ADD             As Long = &H0
Public Const NIM_DELETE          As Long = &H2
Public Const NIF_ICON            As Long = &H2
Public Const NIF_MESSAGE         As Long = &H1
Public Const NIM_MODIFY          As Long = &H1
Public Const NIF_TIP             As Long = &H4
Public Const HWND_NOTOPMOST      As Long = -2
Public Const HWND_TOPMOST        As Long = -1
Public Sub SystrayOn(frm As Form, IconTooltipText As String)
On Error Resume Next
    'adds icon to systray
    vbTray.cbSize = Len(vbTray)
    vbTray.hWnd = frm.hWnd
    vbTray.uId = vbNull
    vbTray.uFlags = NIF_ICON Or NIF_TIP Or NIF_MESSAGE
    vbTray.uCallBackMessage = WM_MOUSEMOVE
    vbTray.szTip = Trim(IconTooltipText$) & vbNullChar
    vbTray.hIcon = frm.Icon
    Call Shell_NotifyIcon(NIM_ADD, vbTray)
    App.TaskVisible = False
End Sub
Public Sub SystrayOff(frm As Form)
On Error Resume Next
    'removes icon from systray
    vbTray.cbSize = Len(vbTray)
    vbTray.hWnd = frm.hWnd
    vbTray.uId = vbNull
    Call Shell_NotifyIcon(NIM_DELETE, vbTray)
End Sub
Public Sub FormOnTop(frm As Form)
On Error Resume Next
    'puts your form ontop of all the other windows!
    Call SetWindowPos(frm.hWnd, HWND_TOPMOST, 0&, 0&, 0&, 0&, FLAGS)
End Sub
