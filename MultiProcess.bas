Attribute VB_Name = "MultiProcess"
'***************************************
'多线程模块还属于实验阶段，不可使用
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
Sub recComm() 'addressof 的函数只能Public？
Dim j As Integer

    Do While 1
        If threadVar = 1 Then
        
            threadVar = 0
        End If
    Loop

End Sub
Sub xywhXC1()
Dim i As Long       '之前代码忘记了声明这个，在线城中是不被允许的。
Dim TempStr As String
     For i = 1 To 10000
         TempStr = CStr(i)
         'Form1.Label1.Caption = TempStr
         Sleep 20
     Next
End Sub
Sub hotKeyprocess()
WaitMessage '等待消息
          If PeekMessage(Message, Form1.hwnd, WM_HOTKEY, WM_HOTKEY, PM_REMOVE) Then '检查是否热键被按下
            End
          End If
         DoEvents '
End Sub
