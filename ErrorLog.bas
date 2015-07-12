Attribute VB_Name = "ErrorLog"
Option Explicit


'使用：
'在Sub开头添加：
'On Error GoTo Errlog
'.......
'Exit Sub
'Errlog:
'Call whenErr(Err.Number, "AppFilename","errsection")

'End Sub

Public Sub whenErr(ByVal num As Integer, ByVal title As String, ByVal ex As String)
Call init_dir(title) '创建log.txt是否包含在程序创建自身文件夹时

Dim tim As String
tim = Year(Date) & " " & Month(Date) & " " & Day(Date)
tim = tim & "  "
tim = tim & Hour(Time) & " " & Minute(Time) & " " & Second(Time) '变量名不要和time,date重复（而且vb不分大小写）

Open "d:\Myuse\" & title & "\" & "log" & ".txt" For Append As #1
Print #1, "Time:" & tim & "  Section:" & ex & "  ErrorNumber:" & CStr(num)
Close #1

Debug.Print "ErrNumber:" & CStr(num)
End Sub

Private Function init_dir(title As String) As Long  '创建本程序文件夹Dir

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
