Attribute VB_Name = "Module3"
Public Declare Function SetWindowPos Lib "user32" (ByVal hwnd As Long, ByVal hWndInsertAfter As Long, ByVal X As Long, ByVal Y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long) As Long
Public Const HWND_TOPMOST = -1
Public Const HWND_NOTOPMOST = -2
Public Const SWP_NOMOVE = &H2
Public Const SWP_NOSIZE = &H1
Public Const SWP_SHOWWINDOW = &H40
Public Const SWP_NOOWNERZORDER = &H200      '  Don't do owner Z ordering
Public Const SWP_NOACTIVATE = &H10
 
Public Declare Function Shell_NotifyIcon Lib "shell32.dll" (ByVal dwMessage As Long, lpData As NOTIFYICONDATA) As Long

Public Const NIM_ADD = &H0                     '在任务栏中增加一个图标
Public Const NIM_DELETE = &H2                  '删除任务栏中的一个图标
Public Const NIM_MODIFY = &H1                  '修改任务栏中个图标信息
Public Const NIF_ICON = &H2                    '
Public Const NIF_MESSAGE = &H1                 'NOTIFYICONDATA结构中uFlags的控制信息
Public Const NIF_TIP = &H4                     '

Public Type NOTIFYICONDATA
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
 
Public changeshape As Integer
Const Rightup = 6
Const updown = 7
Const leftup = 8
Const leftright = 9
Const linetocur = 70
