Attribute VB_Name = "Sim_MouseKey"
Option Explicit

Private Type RECT
        Left As Long
        Top As Long
        Right As Long
        Bottom As Long
End Type

Public Declare Function GetWindowRect Lib "user32" (ByVal hwnd As Long, lpRect As RECT) As Long '��Right���ײ������1������ʵֵ

Private Declare Function SetCursorPos Lib "user32" (ByVal x As Long, ByVal y As Long) As Long

Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (pDst As Any, pSrc As Any, ByVal ByteLen As Long)

'�ڳ�windows��ϵͳ�����������Ӧ�ó��򴰿���ʱҲ����Ч ��Ѳ�����SendInput����ʹ��
Public Declare Sub mouse_event Lib "user32" (ByVal dwFlags As Long, ByVal dx As Long, ByVal dy As Long, ByVal cButtons As Long, ByVal dwExtraInfo As Long)
'�ڳ��ڳ�windows��ϵͳ�����������Ӧ�ó��򴰿�Ҳ����Ч                                             '����type�����ƿ����Լ�����
Private Declare Function SendInput Lib "user32.dll" (ByVal nInputs As Long, pInputs As INPUT_TYPE, ByVal cbSize As Long) As Long
'ע��  Const����  ��Ҫд��
Private Const INPUT_MOUSE = 0
Private Const INPUT_KEYBOARD = 1
Private Const INPUT_HARDWARE = 2
Private Const KEYEVENTF_EXTENDEDKEY = &H1
Private Const KEYEVENTF_KEYUP = &H2
Private Const VK_A = 65

Public Const MOUSEEVENTF_MOVE = &H1 '  mouse move  'ֻ��ConstĬ��Private
Public Const MOUSEEVENTF_LEFTDOWN = &H2 '  left button down
Public Const MOUSEEVENTF_LEFTUP = &H4 '  left button up
Public Const MOUSEEVENTF_RIGHTDOWN = &H8 '  right button down
Public Const MOUSEEVENTF_RIGHTUP = &H10 '  right button up
Public Const MOUSEEVENTF_MIDDLEDOWN = &H20 '  middle button down
Public Const MOUSEEVENTF_MIDDLEUP = &H40 '  middle button up
Public Const MOUSEEVENTF_ABSOLUTE = &H8000 '  absolute move
Public Enum mouseclick  '������곣�� ,Enumֻ��һ��ֻ����Public
 '��������
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

Private inputEvents(0 To 1) As INPUT_TYPE ' �����¼���Ϣ
Private mouseEvent As MOUSEINPUT          '��ʱ�������������Ϣ

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
'Ϊȫ���ȼ����һ����ʶ��
Public Declare Function RegisterHotKey Lib "user32" (ByVal hwnd As Long, ByVal id As Long, _
                ByVal fsModifiers As Long, ByVal vk As Long) As Long
'hWnd�������ȼ�����WM_HOTKEY��Ϣ�Ĵ��ھ��
'id�������ȼ��ı�ʶ��,GlobalAddAtom��������ȼ��ı�ʶ��.
'MOD_ALTΪAlt����MOD_CONTROLΪCtrl����MOD_SHIFTΪShift����MOD_WINΪWindows������
'vk�������ȼ���������롣
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
'mouse_event MOUSEEVENTF_LEFTDOWN, 0, 0, 0, 0                ' ʹ��mouse_evevntAPI����
'Call waittime(0.1)
'mouse_event MOUSEEVENTF_LEFTUP, 0, 0, 0, 0

Public Function SetMousePos(x As Long, y As Long)
    SetCursorPos x, y
End Function

Public Sub VirtualClickMouse(ButtonPressed As mouseclick, Optional ButtonRelease As mouseclick) ' ʹ��SendInput API����

    Dim intX As Integer
   

' Load the information needed to synthesize pressing the mouse button.
    mouseEvent.dx = 0 ' ��ˮƽ�˶�
    mouseEvent.dy = 0 ' ����ֱ�˶�
    mouseEvent.mouseData = 0
    mouseEvent.dwFlags = ButtonPressed ' ��������
    mouseEvent.dwtime = 0 ' ȱʡ
    mouseEvent.dwExtraInfo = 0 ' �Ǳ���
' ���ƽṹ���������黺����
    inputEvents(0).dwType = INPUT_MOUSE ' �������
    CopyMemory inputEvents(0).xi(0), mouseEvent, Len(mouseEvent)
    intX = SendInput(2, inputEvents(0), Len(inputEvents(0))) ''�ڳ�windows�����Ӧ�ó��򴰿�Ҳ����Ч

'�����ɿ��������������ʱ���
Call waittime(0.1)

' ����, �ſ���갴ť��
    mouseEvent.dx = 0
    mouseEvent.dy = 0
    mouseEvent.mouseData = 0
    mouseEvent.dwFlags = ButtonRelease ' ����̧��
    mouseEvent.dwtime = 0
    mouseEvent.dwExtraInfo = 0
    inputEvents(1).dwType = INPUT_MOUSE
    CopyMemory inputEvents(1).xi(0), mouseEvent, Len(mouseEvent)
    intX = SendInput(2, inputEvents(1), Len(inputEvents(1))) '�ڳ�windows�����Ӧ�ó��򴰿�Ҳ����Ч

End Sub

Public Sub mouseclick(ByVal x As Long, ByVal y As Long)
mouse_event MOUSEEVENTF_ABSOLUTE + MOUSEEVENTF_MOVE, x, y, 0, 0 '���ָ����MOUSEEVENTF_ABSOLUTEֵ����dX��dy���б�׼���ľ������꣬��ֵ��0��65535֮�䡣�¼����򽫴�����ӳ�䵽��ʾ���档���꣨0��0��ӳ�䵽��ʾ��������Ͻǣ���65535��65535��ӳ�䵽���½ǡ�
mouse_event MOUSEEVENTF_LEFTDOWN Or MOUSEEVENTF_LEFTUP, 0, 0, 0, 0
End Sub

'************************HotKey Example**************************
'Private HotKey_Flg as Boolean
'Sub Hotkey()
'HotKey_ID = GlobalAddAtom("Ctrl + S")
' If HotKey_Flg = False Then Call RegisterHotKey(Me.hwnd, HotKey_ID, MOD_CONTROL, vbKeyS)    'ע�� Ctrl+ S Ϊ�ȼ�
'End Sub
'Sub unHot()
'if HotKey_Flg = True then Call UnregisterHotKey(Me.hwnd, HotKey_ID)
'End Sub
'Private Sub Timer1_Timer()
'WaitMessage '�ȴ���Ϣ
'          If PeekMessage(Message, Me.hwnd, WM_HOTKEY, WM_HOTKEY, PM_REMOVE) Then '����Ƿ��ȼ�������
'
'          End If
'         DoEvents '
'End Sub
