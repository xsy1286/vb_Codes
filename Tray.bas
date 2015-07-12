Attribute VB_Name = "Tray"
Option Explicit
'//////////////////////////////////////////////////////////////////////////////
'@@summary
'@@require
'@@reference
'@@license
'@@author
'@@create
'@@modify

'//////////////////////////////////////////////////////////////////////////////
'//
'//      ��������
'//
'//////////////////////////////////////////////////////////////////////////////
Public Declare Function Shell_NotifyIcon Lib "shell32.dll" (ByVal dwMessage As Long, lpData As NOTIFYICONDATA) As Long

Public Const NIM_ADD = &H0                     '��������������һ��ͼ��
Public Const NIM_DELETE = &H2                  'ɾ���������е�һ��ͼ��
Public Const NIM_MODIFY = &H1                  '�޸��������и�ͼ����Ϣ
Public Const NIF_ICON = &H2                    '
Public Const NIF_MESSAGE = &H1                 'NOTIFYICONDATA�ṹ��uFlags�Ŀ�����Ϣ
Public Const NIF_TIP = &H4
'
Private Const WM_MOUSEMOVE = &H200

Public Type NOTIFYICONDATA
  cbSize As Long                        '�����ݽṹ�Ĵ�С
  hwnd As Long                          '������������ͼ��Ĵ��ھ��
  uID As Long                           '�������������ͼ��ı�ʶ
  uFlags As Long                        '������ͼ�깦�ܿ��ƣ�����������ֵ����ϣ�һ��ȫ������
                                        'NIF_MESSAGE ��ʾ���Ϳ�����Ϣ��
                                        'NIF_ICON��ʾ��ʾ�������е�ͼ�ꣻ
                                        'NIF_TIP��ʾ�������е�ͼ���ж�̬��ʾ��
  uCallbackMessage As Long '������ͼ��ͨ�������û����򽻻���Ϣ����������Ϣ�Ĵ�����hWnd����
  hIcon As Long '�������е�ͼ��Ŀ��ƾ��
  szTip As String * 64 'ͼ�����ʾ��Ϣ
End Type
Public p_Tray As NOTIFYICONDATA
'------------------------------------------------------------------------------
'       ���г���
'------------------------------------------------------------------------------

'------------------------------------------------------------------------------
'       ������������
'------------------------------------------------------------------------------

'------------------------------------------------------------------------------
'       ���б���
'------------------------------------------------------------------------------

'------------------------------------------------------------------------------
'       ����API
'------------------------------------------------------------------------------

'//////////////////////////////////////////////////////////////////////////////
'//
'//      ˽������
'//
'//////////////////////////////////////////////////////////////////////////////

'------------------------------------------------------------------------------
'       ˽�г���
'------------------------------------------------------------------------------

'------------------------------------------------------------------------------
'       ˽����������
'------------------------------------------------------------------------------

'------------------------------------------------------------------------------
'       ˽�б���
'------------------------------------------------------------------------------

'------------------------------------------------------------------------------
'       ���Ա���
'------------------------------------------------------------------------------

'------------------------------------------------------------------------------
'       ˽��API
'------------------------------------------------------------------------------
'*************************************************************************
'to���̹���
'*************************************************************************
Public Function toTray(frm As Form, sz As String, force As Boolean)
    Debug.Print frm.WindowState

    If force = True Then
        GoTo Tray

    ElseIf frm.WindowState = 1 Then
        GoTo Tray
    Else

        Exit Function

    End If

Tray:
    p_Tray.cbSize = Len(p_Tray)
    p_Tray.uID = vbNull
    p_Tray.hwnd = frm.hwnd
    p_Tray.uFlags = NIF_TIP Or NIF_MESSAGE Or NIF_ICON
    p_Tray.uCallbackMessage = WM_MOUSEMOVE   '���ʹ��Form��Move�¼������ָ�����
    p_Tray.hIcon = frm.Icon
    p_Tray.szTip = sz & vbNullChar
    Shell_NotifyIcon NIM_ADD, p_Tray
    frm.Hide
End Function
'*************************************************************************
'�����ָ����̹���һ������
'*************************************************************************
'Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
'    Dim msg As Long
'    msg = X / 15
'    If msg = WM_LBUTTONDBLCLK Then
'        Me.WindowState = 0
'        Me.Show
'        Shell_NotifyIcon NIM_DELETE, p_Tray
'    End If
'End Sub
'*************************************************************************

'//////////////////////////////////////////////////////////////////////////////
'//
'//      �¼�����
'//
'//////////////////////////////////////////////////////////////////////////////

'//////////////////////////////////////////////////////////////////////////////
'//
'//      ˽������
'//
'//////////////////////////////////////////////////////////////////////////////

'//////////////////////////////////////////////////////////////////////////////
'//
'//      ˽�з���
'//
'//////////////////////////////////////////////////////////////////////////////

'//////////////////////////////////////////////////////////////////////////////
'//
'//      ��������
'//
'//////////////////////////////////////////////////////////////////////////////

'//////////////////////////////////////////////////////////////////////////////
'//
'//      ���з���
'//
'//////////////////////////////////////////////////////////////////////////////
