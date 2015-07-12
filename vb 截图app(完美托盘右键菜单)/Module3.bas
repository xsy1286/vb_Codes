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

Public Const NIM_ADD = &H0                     '��������������һ��ͼ��
Public Const NIM_DELETE = &H2                  'ɾ���������е�һ��ͼ��
Public Const NIM_MODIFY = &H1                  '�޸��������и�ͼ����Ϣ
Public Const NIF_ICON = &H2                    '
Public Const NIF_MESSAGE = &H1                 'NOTIFYICONDATA�ṹ��uFlags�Ŀ�����Ϣ
Public Const NIF_TIP = &H4                     '

Public Type NOTIFYICONDATA
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
 
Public changeshape As Integer
Const Rightup = 6
Const updown = 7
Const leftup = 8
Const leftright = 9
Const linetocur = 70
