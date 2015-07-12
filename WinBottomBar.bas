Attribute VB_Name = "WinBottomBar"
Private Declare Function SystemParametersInfo Lib "user32" Alias "SystemParametersInfoA" (ByVal uAction As Long, ByVal uParam As Long, ByRef lpvParam As Any, ByVal fuWinIni As Long) As Long
Private Const SPI_GETWORKAREA = 48
'Public Type RECT  'vb�ڲ����ж���
'Left As Long
'Top As Long
'Right As Long
'Bottom As Long
'End Type
Dim abc As RECT

Public Declare Function Shell_NotifyIcon Lib "shell32.dll" (ByVal dwMessage As Long, lpData As NOTIFYICONDATA) As Long
Public Const NIM_ADD = &H0                     '��������������һ��ͼ��
Public Const NIM_DELETE = &H2                  'ɾ���������е�һ��ͼ��
Public Const NIM_MODIFY = &H1                  '�޸��������и�ͼ����Ϣ
Public Const NIF_ICON = &H2                    '
Public Const NIF_MESSAGE = &H1                 'NOTIFYICONDATA�ṹ��uFlags�Ŀ�����Ϣ
Public Const NIF_TIP = &H4                     '

Public Type NOTIFYICONDATA   '�ṹ������Public
  cbSize As Long                        '�����ݽṹ�Ĵ�С
  hwnd As Long                          '������������ͼ��Ĵ��ھ��
  uID As Long                           '�������������ͼ��ı�ʶ
  uFlags As Long                        '������ͼ�깦�ܿ��ƣ�����������ֵ����ϣ�һ��ȫ������
                                        'NIF_MESSAGE ��ʾ���Ϳ�����Ϣ��
                                        'NIF_ICON��ʾ��ʾ�������е�ͼ�ꣻ
                                        'NIF_TIP��ʾ�������е�ͼ���ж�̬��ʾ��
  uCallbackMessage As Long '������ͼ��ͨ�������û����򽻻���Ϣ���������Ϣ�Ĵ�����hWnd����
  hIcon As Long '�������е�ͼ��Ŀ��ƾ��
  szTip As String * 64 'ͼ�����ʾ��Ϣ
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
Public Function minTobar(Frm As Form, nam As String, ByRef Tray As NOTIFYICONDATA) 'ע������Լ���Form�ڶ���һ��trayN

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
'**********************��С�������÷�********************
' '��Form_Resize�е���ʱ
'Private Sub Form_Resize()
'
'    If Me.WindowState = 1 Then
'        Call minTobar(Me, "name", tray1)
'    End If
'
'End Sub
'
'************���˫���ָ�-Mouse.Button********
'Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
'
'    If Button = vbLeftButton Then
'            Me.WindowState = 0
'            Me.Show
'            Shell_NotifyIcon NIM_DELETE, tray1  'ȡ������
'
'    End If
'End Sub

'    '��˫������ʱ�ָ�ԭ״-winMsg
'Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single) '��FormnameΪ����Form������
'
'        Dim Msg As Long
'
'        Msg = X '/ Screen.TwipsPerPixelX   '���˾�ԭ��
'
'        If Msg = WM_LBUTTONDBLCLK Then
'            Me.WindowState = 0
'            Me.Show
'            Shell_NotifyIcon NIM_DELETE, trayN   'ȡ������
'
'        ElseIf Msg = WM_RBUTTONDOWN Then   '����ʱ�Ҽ�
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
