Attribute VB_Name = "MultiProcess"
'***************************************
'���߳�ģ�黹����ʵ��׶Σ�����ʹ��
'***************************************
Option Explicit
Public Declare Function CreateThread Lib "kernel32" (lpThreadAttributes As Any, ByVal dwStackSize As Long, ByVal lpStartAddress As Long, ByVal lpParameter As Long, ByVal dwCreationFlags As Long, lpThreadId As Long) As Long
Public Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)
Public Declare Function TerminateThread Lib "kernel32" (ByVal hThread As Long, ByVal dwExitCode As Long) As Long
Public id As Long
Public threadVar As Long
Dim BytReceived() As Byte
Dim strBuff As String
'Usage:
 'id = CreateThread(ByVal 0&, ByVal 0&, AddressOf AddText, ByVal 0&, 0, id)
 'Call TerminateThread(id, ByVal 0&)
Sub recComm() 'addressof �ĺ���ֻ��Public��
Dim j As Integer

    Do While 1
        If threadVar = 1 Then
        
            threadVar = 0
        End If
    Loop

End Sub
Sub xywhXC1()
Dim i As Long       '֮ǰ����������������������߳����ǲ�������ġ�
Dim TempStr As String
     For i = 1 To 10000
         TempStr = CStr(i)
         'Form1.Label1.Caption = TempStr
         Sleep 20
     Next
End Sub
Sub hotKeyprocess()
WaitMessage '�ȴ���Ϣ
          If PeekMessage(Message, Form1.hwnd, WM_HOTKEY, WM_HOTKEY, PM_REMOVE) Then '����Ƿ��ȼ�������
            End
          End If
         DoEvents '
End Sub
