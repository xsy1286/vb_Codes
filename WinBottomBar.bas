Attribute VB_Name = "WinBottomBar"
Private Declare Function SystemParametersInfo Lib "user32" Alias "SystemParametersInfoA" (ByVal uAction As Long, ByVal uParam As Long, ByRef lpvParam As Any, ByVal fuWinIni As Long) As Long
Private Const SPI_GETWORKAREA = 48
'Public Type RECT  'vb内部已有定义
'Left As Long
'Top As Long
'Right As Long
'Bottom As Long
'End Type
Dim abc As RECT

Public Declare Function Shell_NotifyIcon Lib "shell32.dll" (ByVal dwMessage As Long, lpData As NOTIFYICONDATA) As Long
Public Const NIM_ADD = &H0                     '在任务栏中增加一个图标
Public Const NIM_DELETE = &H2                  '删除任务栏中的一个图标
Public Const NIM_MODIFY = &H1                  '修改任务栏中个图标信息
Public Const NIF_ICON = &H2                    '
Public Const NIF_MESSAGE = &H1                 'NOTIFYICONDATA结构中uFlags的控制信息
Public Const NIF_TIP = &H4                     '

Public Type NOTIFYICONDATA   '结构基本上Public
  cbSize As Long                        '该数据结构的大小
  hwnd As Long                          '处理任务栏中图标的窗口句柄
  uID As Long                           '定义的任务栏中图标的标识
  uFlags As Long                        '任务栏图标功能控制，可以是以下值的组合（一般全包括）
                                        'NIF_MESSAGE 表示发送控制消息；
                                        'NIF_ICON表示显示控制栏中的图标；
                                        'NIF_TIP表示任务栏中的图标有动态提示。
  uCallbackMessage As Long '任务栏图标通过它与用户程序交换消息，处理该消息的窗口由hWnd决定
  hIcon As Long '任务栏中的图标的控制句柄
  szTip As String * 64 '图标的提示信息
End Type

Public Const WM_MOUSEMOVE = &H200
Public Const WM_LBUTTONDOWN = &H201
Public Const WM_LBUTTONUP = &H202
Public Const WM_LBUTTONDBLCLK = &H203
Public Const WM_RBUTTONDOWN = &H204
Public Const WM_RBUTTONUP = &H205
Public Const WM_RBUTTONDBLCLK = &H206
Public Const WM_MBUTTONDOWN = &H207
Public Const WM_MBUTTONUP = &H208
Public Const WM_MBUTTONDBLCLK = &H209
Public Const WM_MOUSEACTIVATE = &H21
Public Const WM_MOUSEFIRST = &H200
Public Const WM_MOUSELAST = &H209
Public Const WM_MOUSEWHEEL = &H20A
'Private trayX As NOTIFYICONDATA
Public Function minTobar(Frm As Form, nam As String, ByRef Tray As NOTIFYICONDATA) '注意得在自己的Form内定义一个trayN

    Tray.cbSize = Len(Tray)
    Tray.uID = vbNull
    Tray.hwnd = Frm.hwnd
    Tray.uFlags = NIF_TIP Or NIF_MESSAGE Or NIF_ICON
    Tray.uCallbackMessage = WM_MOUSEMOVE
    Tray.hIcon = Frm.Icon
    Tray.szTip = nam & vbNullChar
    Shell_NotifyIcon NIM_ADD, Tray
    Frm.Hide

End Function
'**********************缩小到托盘用法********************
' '在Form_Resize中调用时
'Private Sub Form_Resize()
'
'    If Me.WindowState = 1 Then
'        Call minTobar(Me, "name", tray1)
'    End If
'
'End Sub
'
'************左键双击恢复-Mouse.Button********
'Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
'
'    If Button = vbLeftButton Then
'            Me.WindowState = 0
'            Me.Show
'            Shell_NotifyIcon NIM_DELETE, tray1  '取消托盘
'
'    End If
'End Sub

'    '当双击托盘时恢复原状-winMsg
'Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single) '改Formname为所用Form的名称
'
'        Dim Msg As Long
'
'        Msg = X '/ Screen.TwipsPerPixelX   '？此句原理
'
'        If Msg = WM_LBUTTONDBLCLK Then
'            Me.WindowState = 0
'            Me.Show
'            Shell_NotifyIcon NIM_DELETE, trayN   '取消托盘
'
'        ElseIf Msg = WM_RBUTTONDOWN Then   '托盘时右键
'
'            Dim p As POINTAPI
'
'            Call GetCursorPos(p)
'
'        End If
'End Sub


Public Function GetTaskbar(rectVal As RECT) As Long

GetTaskbar = SystemParametersInfo(SPI_GETWORKAREA, 0, rectVal, 0)

End Function
