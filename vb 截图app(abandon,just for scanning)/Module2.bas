Attribute VB_Name = "Module2"
Option Explicit
Private Type POINTAPI
    x As Long
    y As Long
End Type
Public Type Msg
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

Public HotKey_ID As Long
Public HotKey_Flg As Boolean
Public Message As Msg
Public Declare Function GetLastError Lib "kernel32" () As Long
Public Declare Function GlobalAddAtom Lib "kernel32" Alias "GlobalAddAtomA" (ByVal lpString _
                As String) As Integer
'Ϊȫ���ȼ����һ����ʶ��
Public Declare Function RegisterHotKey Lib "user32" (ByVal hwnd As Long, ByVal ID As Long, _
                ByVal fsModifiers As Long, ByVal vk As Long) As Long
'hWnd�������ȼ�����WM_HOTKEY��Ϣ�Ĵ��ھ��
'id�������ȼ��ı�ʶ��,GlobalAddAtom��������ȼ��ı�ʶ��.
'MOD_ALTΪAlt����MOD_CONTROLΪCtrl����MOD_SHIFTΪShift����MOD_WINΪWindows������
'vk�������ȼ���������롣
Public Declare Function UnregisterHotKey Lib "user32" (ByVal hwnd As Long, ByVal ID As Long _
                ) As Long
                
Public Declare Function PeekMessage Lib "user32" Alias "PeekMessageA" (lpMsg As Msg, ByVal _
                hwnd As Long, ByVal wMsgFilterMin As Long, ByVal wMsgFilterMax As Long, ByVal _
                wRemoveMsg As Long) As Long
Public Declare Function WaitMessage Lib "user32" () As Long
Public Declare Function GetTickCount Lib "kernel32" () As Long
Public v, v4, v5, v6 As Integer
Public dn As Integer  'boolean  Ҳ��
Public tx, ty As Single

Public Sub delaySecond(ByVal n As Single)   '���� delay ������ʱ����ʱ���� n ���� Single��
Dim tm1 As Single, tm2 As Single  '�����¼ʱ����� tm1  tm2  ���� Single
tm1 = Timer  '��¼ϵͳ����ʱ�䵽 tm1  ��ʱ��Ϊ��ʼ��ʱʱ��
Do  'ѭ���� ��ִ�е�Loop �����������
tm2 = Timer  '��ֵ����ʱ�䵽 tm2  ��ʱ��Ϊϵͳ��ǰʱ�䣬����ʱ��ʼʱ��֮��
If tm2 < tm1 Then tm2 = tm2 + 86400  '�����ж�������� ������˵��#1
If tm2 - tm1 > n Then Exit Do  '�ж�tm2-tm1������ʱ������n��������ѭ������˵��#2
DoEvents     'ת�ÿ���Ȩ���Ա��ò���ϵͳ�����������¼�

Loop  '����ѭ���� ����ѭ��
End Sub

Public Sub delay(numa As Long)

    Dim num1 As Long
  Dim num2 As Long
  Dim numb As Long
  numb = 0
  num1 = GetTickCount
    Do While numa - numb > 0
    num2 = GetTickCount
  numb = num2 - num1
DoEvents
    Loop
  End Sub
  Public Sub Cancelshoot()

Call UnHooK
If v = 0 Then
Form1.Show
Form1.Refresh
Form1.Check1.Refresh
Form1.Check2.Refresh
Form1.Check3.Refresh
Form1.Check4.Refresh
Form1.Command1.Refresh
Form1.Label1.Refresh
Form1.Label2.Refresh
Form1.Label3.Refresh
Form1.Text4.Refresh
Form1.Combo1.Refresh
End If
 Unload Form2

  End Sub
  Public Sub Down()
  Form1.Picture1.Refresh
Form2.Text1.Text = Year(Date) & Month(Date) & Day(Date) & Hour(time) & Minute(time) & Second(time) & "��ͼ"
 Call UnHooK
Form2.Shape1.Visible = False
Form2.Shape1.Width = 0
Form2.Shape1.Height = 0
Form2.Command3.Visible = False
Form2.Command4.Visible = False
Form1.Picture1.Refresh

If v4 = 1 Then
SavePicture Form1.Picture1.Image, (Form1.Text4.Text & "\" & Form2.Text1.Text & Form1.Combo1.Text)
If v = 0 Then
Form1.Show
Form1.Refresh
Form1.Check1.Refresh
Form1.Check2.Refresh
Form1.Check3.Refresh
Form1.Check4.Refresh
Form1.Command1.Refresh
Form1.Label1.Refresh
Form1.Label2.Refresh
Form1.Label3.Refresh
Form1.Text4.Refresh
Form1.Combo1.Refresh
End If

Form2.Hide

 
                 If Form2.Check1.Value = 1 Then Shell "explorer " & Form2.CommonDialog1.FileName, 1  'shell��������C��  Shell "explorer ·��"
 
 Unload Form2

Else:
Form2.CommonDialog1.ShowSave

'i
If Len(Form2.CommonDialog1.FileName) <> 0 Then
SavePicture Form1.Picture1.Image, (Form2.CommonDialog1.FileName)

If v = 0 Then
Form1.Show
Form1.Refresh
Form1.Check1.Refresh
Form1.Check2.Refresh
Form1.Check3.Refresh
Form1.Check4.Refresh
Form1.Command1.Refresh
Form1.Label1.Refresh
Form1.Label2.Refresh
Form1.Label3.Refresh
Form1.Text4.Refresh
Form1.Combo1.Refresh
End If

Form2.Hide

 
                 If Form2.Check1.Value = 1 Then Shell "explorer " & Form2.CommonDialog1.FileName, 1  'shell��������C��  Shell "explorer ·��"
 
 Unload Form2
Else:
Call HooK
End If
'i

End If
 

  End Sub



Public Sub dl()
If v5 = 1 Then
Form1.Picture1.Refresh
Clipboard.SetData Form1.Picture1.Image
End If

Call Down

End Sub
