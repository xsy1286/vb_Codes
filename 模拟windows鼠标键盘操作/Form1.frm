VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   2340
   ClientLeft      =   105
   ClientTop       =   435
   ClientWidth     =   3630
   LinkTopic       =   "Form1"
   ScaleHeight     =   2340
   ScaleWidth      =   3630
   StartUpPosition =   3  '窗口缺省
   Begin VB.Timer Timer1 
      Left            =   2160
      Top             =   1560
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Command2"
      Height          =   492
      Left            =   240
      TabIndex        =   1
      Top             =   1320
      Width           =   972
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Command1"
      Height          =   612
      Left            =   1080
      TabIndex        =   0
      Top             =   720
      Width           =   2172
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Declare Function SetCursorPos Lib "user32" (ByVal x As Long, ByVal y As Long) As Long

Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (pDst As Any, pSrc As Any, ByVal ByteLen As Long)

'在除windows子系统自身窗口以外的应用程序窗口有时也能有效 最佳操作与SendInput合作使用
Private Declare Sub mouse_event Lib "user32" (ByVal dwFlags As Long, ByVal dx As Long, ByVal dy As Long, ByVal cButtons As Long, ByVal dwExtraInfo As Long)
'在除windows子系统自身窗口以外的应用程序窗口也能有效                                             '以下type的名称可以自己定义
Private Declare Function SendInput Lib "user32.dll" (ByVal nInputs As Long, pInputs As INPUT_TYPE, ByVal cbSize As Long) As Long

'PsotMessage 需知窗口句柄
Private Declare Function PostMessage Lib "user32" Alias "PostMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long

'包括windows界面及应用程序界面都可以输入键，为输入物理按键为bVK
Private Declare Sub keybd_event Lib "user32" (ByVal bVk As Byte, ByVal bScan As Byte, ByVal dwFlags As Long, ByVal dwExtraInfo As Long)

Private Declare Function MapVirtualKey Lib "user32" Alias "MapVirtualKeyA" (ByVal wCode As Long, ByVal wMapType As Long) As Long
Private Declare Function SetKeyboardState Lib "user32" (lppbKeyState As Byte) As Long
Private Declare Function GetKeyState Lib "user32" (ByVal nVirtKey As Long) As Integer

'注意  Const常数  都要写好
Private Const INPUT_MOUSE = 0
Private Const INPUT_KEYBOARD = 1
Private Const INPUT_HARDWARE = 2
Private Const KEYEVENTF_EXTENDEDKEY = &H1
Private Const KEYEVENTF_KEYUP = &H2
Private Const VK_A = 65
Private Enum MouseClick '定义鼠标常数
    MOUSEEVENTF_LEFTDOWN = &H2
    MOUSEEVENTF_LEFTUP = &H4
    MOUSEEVENTF_RIGHTDOWN = &H8
    MOUSEEVENTF_RIGHTUP = &H10
    MOUSEEVENTF_MIDDLEDOWN = &H20
    MOUSEEVENTF_MIDDLEUP = &H40
End Enum

 Private Type MOUSEINPUT
    dx As Long
    dy As Long
    mouseData As Long
    dwFlags As Long
    dwtime As Long
    dwExtraInfo As Long
End Type

Private Type KEYBDINPUT
    wVk As Integer
    wScan As Integer
    dwFlags As Long
    time As Long
    dwExtraInfo As Long
End Type
Private Type INPUT_TYPE
    dwType As Long
    xi(0 To 23) As Byte
End Type

 Dim inputEvents(0 To 1) As INPUT_TYPE ' 锁定事件信息
    Dim mouseEvent As MOUSEINPUT          '临时锁定鼠标输入信息
Dim GInput(0 To 1) As INPUT_TYPE
Dim KInput As KEYBDINPUT

Private Function waittime(delay As Single)
  Dim starttime As Single
  starttime = Timer
  Do Until (Timer - starttime) > delay
  DoEvents
  Loop
End Function
Private Sub Command1_Click()
'Shell "taskkill /f /im explorer.exe"

End Sub
Private Function AutoMouse(x As Long, y As Long)
    SetCursorPos x, y

End Function
Private Function MousePress()
 mouse_event MOUSEEVENTF_LEFTDOWN Or MOUSEEVENTF_LEFTUP, 0, 0, 0, 0

End Function

Private Sub VirtualClickMouse(ButtonPressed As MouseClick, Optional ButtonRelease As MouseClick)

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

Private Sub Command2_Click()
Call waittime(3)


VirtualClickMouse MOUSEEVENTF_LEFTDOWN, MOUSEEVENTF_LEFTUP

Timer1.Interval = 500
End Sub

Private Sub Form_Click()
Shell "taskkill /f /im Porsche.exe"
Call waittime(1)
 'vb打开快捷方式 可解决直接Shell问题及打开直接打开需参数问题
If Dir("C:\Users\Administrator\Desktop\Porsche.lnk") <> "" Then Shell "Rundll32 url.dll, FileProtocolHandler C:\Users\Administrator\Desktop\Porsche.lnk"


Call waittime(14)
Call AutoMouse(320, 394)
Call waittime(2)
VirtualClickMouse MOUSEEVENTF_LEFTDOWN, MOUSEEVENTF_LEFTUP


Call waittime(3.5)
Call AutoMouse(223, 60)
Call waittime(3)
VirtualClickMouse MOUSEEVENTF_LEFTDOWN, MOUSEEVENTF_LEFTUP
Call waittime(0.3)
mouse_event MOUSEEVENTF_LEFTDOWN, 0, 0, 0, 0
Call waittime(0.1)
mouse_event MOUSEEVENTF_LEFTUP, 0, 0, 0, 0
   
Call waittime(3)
Call AutoMouse(274, 268)
Call waittime(2)
VirtualClickMouse MOUSEEVENTF_LEFTDOWN, MOUSEEVENTF_LEFTUP
Call waittime(0.3)
mouse_event MOUSEEVENTF_LEFTDOWN, 0, 0, 0, 0
Call waittime(0.1)
mouse_event MOUSEEVENTF_LEFTUP, 0, 0, 0, 0

Call waittime(3)
Call AutoMouse(345, 472)
Call waittime(2)
VirtualClickMouse MOUSEEVENTF_LEFTDOWN, MOUSEEVENTF_LEFTUP
Call waittime(0.3)
mouse_event MOUSEEVENTF_LEFTDOWN, 0, 0, 0, 0
Call waittime(0.1)
mouse_event MOUSEEVENTF_LEFTUP, 0, 0, 0, 0

Form1.Hide
Unload Form1
End Sub

'模拟键盘会被360报错
Private Sub T4A() '为A键
keybd_event VK_A, 0, 0, 0 'keydown
keybd_event VK_A, 0, KEYEVENTF_KEYUP, 0
End Sub

Private Sub T5A()
Dim ret As Long
KInput.wVk = VK_A
KInput.dwFlags = 0
GInput(0).dwType = INPUT_KEYBOARD
CopyMemory GInput(0).xi(0), KInput, Len(KInput)
ret = SendInput(2, GInput(0), Len(GInput(0)))

'以下为UP
KInput.wVk = VK_A
KInput.dwFlags = KEYEVENTF_KEYUP
GInput(1).dwType = INPUT_KEYBOARD
CopyMemory GInput(1).xi(0), KInput, Len(KInput)
ret = SendInput(2, GInput(1), Len(GInput(1)))
End Sub

Private Sub Timer1_Timer()
Call waittime(0.3)
mouse_event MOUSEEVENTF_LEFTDOWN, 0, 0, 0, 0
Call waittime(0.1)
mouse_event MOUSEEVENTF_LEFTUP, 0, 0, 0, 0
End Sub
