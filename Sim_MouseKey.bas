Attribute VB_Name = "Sim_MouseKey"
Option Explicit

Private Type RECT
        Left As Long
        Top As Long
        Right As Long
        Bottom As Long
End Type

Public Declare Function GetWindowRect Lib "user32" (ByVal hwnd As Long, lpRect As RECT) As Long '，Right跟底部都需减1才是真实值

Private Declare Function SetCursorPos Lib "user32" (ByVal x As Long, ByVal y As Long) As Long

Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (pDst As Any, pSrc As Any, ByVal ByteLen As Long)

'在除windows子系统自身窗口以外的应用程序窗口有时也能有效 最佳操作与SendInput合作使用
Public Declare Sub mouse_event Lib "user32" (ByVal dwFlags As Long, ByVal dx As Long, ByVal dy As Long, ByVal cButtons As Long, ByVal dwExtraInfo As Long)
'在除在除windows子系统自身窗口以外的应用程序窗口也能有效                                             '以下type的名称可以自己定义
Private Declare Function SendInput Lib "user32.dll" (ByVal nInputs As Long, pInputs As INPUT_TYPE, ByVal cbSize As Long) As Long
'注意  Const常数  都要写好
Private Const INPUT_MOUSE = 0
Private Const INPUT_KEYBOARD = 1
Private Const INPUT_HARDWARE = 2
Private Const KEYEVENTF_EXTENDEDKEY = &H1
Private Const KEYEVENTF_KEYUP = &H2
Private Const VK_A = 65

Public Const MOUSEEVENTF_MOVE = &H1 '  mouse move  '只有Const默认Private
Public Const MOUSEEVENTF_LEFTDOWN = &H2 '  left button down
Public Const MOUSEEVENTF_LEFTUP = &H4 '  left button up
Public Const MOUSEEVENTF_RIGHTDOWN = &H8 '  right button down
Public Const MOUSEEVENTF_RIGHTUP = &H10 '  right button up
Public Const MOUSEEVENTF_MIDDLEDOWN = &H20 '  middle button down
Public Const MOUSEEVENTF_MIDDLEUP = &H40 '  middle button up
Public Const MOUSEEVENTF_ABSOLUTE = &H8000 '  absolute move
Public Enum mouseclick  '定义鼠标常数 ,Enum只能一般只能做Public
 '上面已有
 tmp = 0&
End Enum

Private Type POINTAPI
        x As Long
        y As Long
End Type

 Public Type MOUSEINPUT
    dx As Long
    dy As Long
    mouseData As Long
    dwFlags As Long
    dwtime As Long
    dwExtraInfo As Long
End Type
Public Type INPUT_TYPE
    dwType As Long
    xi(0 To 23) As Byte
End Type

Private inputEvents(0 To 1) As INPUT_TYPE ' 锁定事件信息
Private mouseEvent As MOUSEINPUT          '临时锁定鼠标输入信息

Private Type Msg
        hwnd As Long
        Message As Long
        wParam As Long
        lParam As Long
        time As Long
        pt As POINTAPI
End Type

Public Const MOD_ALT = &H1
Public Const MOD_CONTROL = &H2
Public Const MOD_SHIFT = &H4
Public Const PM_REMOVE = &H1
Public Const WM_HOTKEY = &H312

'Public HotKey_ID As Long
'Public HotKey_Flg As Boolean
Public Message As Msg
Public Declare Function PeekMessage Lib "user32" Alias "PeekMessageA" (lpMsg As Msg, ByVal _
                hwnd As Long, ByVal wMsgFilterMin As Long, ByVal wMsgFilterMax As Long, ByVal _
                wRemoveMsg As Long) As Long
Public Declare Function WaitMessage Lib "user32" () As Long
Public Declare Function GlobalAddAtom Lib "kernel32" Alias "GlobalAddAtomA" (ByVal lpString _
                As String) As Integer
'为全局热键添加一个标识符
Public Declare Function RegisterHotKey Lib "user32" (ByVal hwnd As Long, ByVal id As Long, _
                ByVal fsModifiers As Long, ByVal vk As Long) As Long
'hWnd：接收热键产生WM_HOTKEY消息的窗口句柄
'id：定义热键的标识符,GlobalAddAtom函数获得热键的标识符.
'MOD_ALT为Alt键，MOD_CONTROL为Ctrl键，MOD_SHIFT为Shift键，MOD_WIN为Windows按键。
'vk：定义热键的虚拟键码。
Public Declare Function UnregisterHotKey Lib "user32" (ByVal hwnd As Long, ByVal id As Long _
                ) As Long
Public Function waittime(delay As Single)
  Dim starttime As Single
  starttime = Timer
  Do Until (Timer - starttime) > delay
  DoEvents
  Loop
End Function


'Example:
'Call waittime(3.5)
'Call SetMousePos(223, 60)
'Call waittime(3)
'VirtualClickMouse MOUSEEVENTF_LEFTDOWN, MOUSEEVENTF_LEFTUP
'Call waittime(0.3)
'mouse_event MOUSEEVENTF_LEFTDOWN, 0, 0, 0, 0                ' 使用mouse_evevntAPI函数
'Call waittime(0.1)
'mouse_event MOUSEEVENTF_LEFTUP, 0, 0, 0, 0

Public Function SetMousePos(x As Long, y As Long)
    SetCursorPos x, y
End Function

Public Sub VirtualClickMouse(ButtonPressed As mouseclick, Optional ButtonRelease As mouseclick) ' 使用SendInput API函数

    Dim intX As Integer
   

' Load the information needed to synthesize pressing the mouse button.
    mouseEvent.dx = 0 ' 不水平运动
    mouseEvent.dy = 0 ' 不垂直运动
    mouseEvent.mouseData = 0
    mouseEvent.dwFlags = ButtonPressed ' 按键按下
    mouseEvent.dwtime = 0 ' 缺省
    mouseEvent.dwExtraInfo = 0 ' 非必须
' 复制结构到输入数组缓冲区
    inputEvents(0).dwType = INPUT_MOUSE ' 鼠标输入
    CopyMemory inputEvents(0).xi(0), mouseEvent, Len(mouseEvent)
    intX = SendInput(2, inputEvents(0), Len(inputEvents(0))) ''在除windows以外的应用程序窗口也能有效

'按下松开鼠标键间必须有延时间隔
Call waittime(0.1)

' 相上, 放开鼠标按钮。
    mouseEvent.dx = 0
    mouseEvent.dy = 0
    mouseEvent.mouseData = 0
    mouseEvent.dwFlags = ButtonRelease ' 按键抬起
    mouseEvent.dwtime = 0
    mouseEvent.dwExtraInfo = 0
    inputEvents(1).dwType = INPUT_MOUSE
    CopyMemory inputEvents(1).xi(0), mouseEvent, Len(mouseEvent)
    intX = SendInput(2, inputEvents(1), Len(inputEvents(1))) '在除windows以外的应用程序窗口也能有效

End Sub

Public Sub mouseclick(ByVal x As Long, ByVal y As Long)
mouse_event MOUSEEVENTF_ABSOLUTE + MOUSEEVENTF_MOVE, x, y, 0, 0 '如果指定了MOUSEEVENTF_ABSOLUTE值，则dX和dy含有标准化的绝对坐标，其值在0到65535之间。事件程序将此坐标映射到显示表面。坐标（0，0）映射到显示表面的左上角，（65535，65535）映射到右下角。
mouse_event MOUSEEVENTF_LEFTDOWN Or MOUSEEVENTF_LEFTUP, 0, 0, 0, 0
End Sub

'************************HotKey Example**************************
'Private HotKey_Flg as Boolean
'Sub Hotkey()
'HotKey_ID = GlobalAddAtom("Ctrl + S")
' If HotKey_Flg = False Then Call RegisterHotKey(Me.hwnd, HotKey_ID, MOD_CONTROL, vbKeyS)    '注册 Ctrl+ S 为热键
'End Sub
'Sub unHot()
'if HotKey_Flg = True then Call UnregisterHotKey(Me.hwnd, HotKey_ID)
'End Sub
'Private Sub Timer1_Timer()
'WaitMessage '等待消息
'          If PeekMessage(Message, Me.hwnd, WM_HOTKEY, WM_HOTKEY, PM_REMOVE) Then '检查是否热键被按下
'
'          End If
'         DoEvents '
'End Sub
