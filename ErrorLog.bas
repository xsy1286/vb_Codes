Attribute VB_Name = "ErrorLog"
Option Explicit


'ʹ�ã�
'��Sub��ͷ��ӣ�
'On Error GoTo Errlog
'.......
'Exit Sub
'Errlog:
'Call whenErr(Err.Number, "AppFilename","errsection")

'End Sub

Public Sub whenErr(ByVal num As Integer, ByVal title As String, ByVal ex As String)
Call init_dir(title) '����log.txt�Ƿ�����ڳ��򴴽������ļ���ʱ

Dim tim As String
tim = Year(Date) & " " & Month(Date) & " " & Day(Date)
tim = tim & "  "
tim = tim & Hour(Time) & " " & Minute(Time) & " " & Second(Time) '��������Ҫ��time,date�ظ�������vb���ִ�Сд��

Open "d:\Myuse\" & title & "\" & "log" & ".txt" For Append As #1
Print #1, "Time:" & tim & "  Section:" & ex & "  ErrorNumber:" & CStr(num)
Close #1

Debug.Print "ErrNumber:" & CStr(num)
End Sub

Private Function init_dir(title As String) As Long  '�����������ļ���Dir

   If Dir("D:\Myuse", vbDirectory) = "" Then
   MkDir ("D:\Myuse")
   End If
   
    If Dir("D:\Myuse\" & title, vbDirectory) = "" Then
    MkDir ("D:\Myuse\" & title)
    End If
    
End Function

Public Sub Ldbg(inp As Long)
Debug.Print CStr(inp)
End Sub
Public Sub Sdbg(inp As String)
Debug.Print inp
End Sub
