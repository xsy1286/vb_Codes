Attribute VB_Name = "HotKey"
Option Explicit
Private Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" _
(ByVal hwnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
Private Declare Function GetWindowLong Lib "user32" Alias "GetWindowLongA" _
(ByVal hwnd As Long, ByVal nIndex As Long) As Long
Private Declare Function CallWindowProc Lib "user32" Alias "CallWindowProcA" _
(ByVal lpPrevWndFunc As Long, ByVal hwnd As Long, ByVal Msg As Long, _
  ByVal wParam As Long, ByVal lParam As Long) As Long
Private Const GWL_WNDPROC = (-4)
'---------------------------------------------------
'ȫ���ȼ�����constants
Private Const MOD_ALT = &H1
Private Const MOD_CONTROL = &H2
Private Const MOD_SHIFT = &H4
Private Const PM_REMOVE = &H1
Private Const WM_HOTKEY = &H312
'---------------------------------------------------
'ȫ���ȼ����蹫������
Private Type POINTAPI
    X As Long
    Y As Long
End Type
Private Type Msg
        hwnd As Long
        Message As Long
        wParam As Long
        lParam As Long
        time As Long
        pt As POINTAPI
End Type
 'Public Message As Msg'�ͱ��.bas ��ͻ


Private Modifiers(0 To 9) As Long, uVirtKey(0 To 9) As Long, idHotKey(0 To 9) As Long



Dim number As Integer
'----------------------------------------------------
'Ϊȫ���ȼ����һ����ʶ��


'Public HotKey_Flg As Boolean

Private Declare Function RegisterHotKey Lib "user32" (ByVal hwnd As Long, ByVal ID As Long, _
                ByVal fsModifiers As Long, ByVal vk As Long) As Long
'hWnd�������ȼ�����WM_HOTKEY��Ϣ�Ĵ��ھ��
'id�������ȼ��ı�ʶ��,GlobalAddAtom��������ȼ��ı�ʶ��.
'MOD_ALTΪAlt����MOD_CONTROLΪCtrl����MOD_SHIFTΪShift����MOD_WINΪWindows������
'vk�������ȼ���������롣
Private Declare Function UnregisterHotKey Lib "user32" (ByVal hwnd As Long, ByVal ID As Long _
                ) As Long
                
'Public Declare Function PeekMessage Lib "user32" Alias "PeekMessageA" (lpMsg As Msg, ByVal _   '�ͱ��.bas ��ͻ
'                hwnd As Long, ByVal wMsgFilterMin As Long, ByVal wMsgFilterMax As Long, ByVal _
'                wRemoveMsg As Long) As Long
Public Declare Function WaitMessage Lib "user32" () As Long



'----------------------------------------------------
'Example: 'ȱ�㣺Timer�Ļ�©����ʱ��HotKey
'Private Sub Timer1_Timer()
'WaitMessage '�ȴ���Ϣ
'          If PeekMessage(Message, Me.hwnd, WM_HOTKEY, WM_HOTKEY, PM_REMOVE) Then '����Ƿ��ȼ�������
'             'Todo,when Hotkey is pressed
'
'          End If
'
'         DoEvents 'ת�ÿ���Ȩ,�������ϵͳ���������¼�
'
'End Sub

'Ŀǰͬʱ��ֻ��ע��һ���ȼ�,num���ֲ�����number��������
Public Function insertHotKey(ByVal hwd As Long, ByRef HotKey_ID As Long, ByVal ctrl As Boolean, _
                                            ByVal alt As Boolean, ByVal shf As Boolean, ByVal key As Integer, ByRef num As Integer) As Long
  Dim s As String, ret&
  If number > 9 Then MsgBox "HotKey Full": insertHotKey = 0: Exit Function
  
  If ctrl = False And shf = False And alt = False Then insertHotKey = 0: Exit Function
  Modifiers(number) = 0
  idHotKey(number) = 0
  uVirtKey(number) = 0
  
  If ctrl = True Then s = "Ctrl +": Modifiers(number) = Modifiers(number) + MOD_CONTROL
  If alt = True Then s = s + "Alt +": Modifiers(number) = Modifiers(number) + MOD_ALT
  If shf = True Then s = s + "Shift +": Modifiers(number) = Modifiers(number) + MOD_SHIFT
  s = s + Chr(key)
  
    HotKey_ID = GlobalAddAtom(s)
  If HotKey_ID = 0 Then insertHotKey = 0: Exit Function
 
  
'If number = 0 Then
 
 ret = RegisterHotKey(hwd, HotKey_ID, Modifiers(number), key)
 If ret = 0 Then insertHotKey = 0: Exit Function

 idHotKey(number) = HotKey_ID
  uVirtKey(number) = key
   
num = number
insertHotKey = 1: 'number = number + 1
End Function


Public Function unHotKey(ByVal hwd As Long, ByRef HotKey_ID As Long) As Long

    Dim ret As Long


    ret = UnregisterHotKey(hwd, HotKey_ID)

    If ret = 0 Then unHotKey = 0: Exit Function

'number=number-1 '����������ȼ�ʧȥ�����м��ȼ���ȱ
End Function
