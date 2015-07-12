Attribute VB_Name = "txt_init"
Option Explicit
Public Function init_dir(title As String) As Long  '创建本程序文件夹Dir

   If Dir("D:\Myuse", vbDirectory) = "" Then
   MkDir ("D:\Myuse")
   End If
   
    If Dir("D:\Myuse\" & title, vbDirectory) = "" Then
    MkDir ("D:\Myuse\" & title)
    End If
    
End Function

Public Function init_txt(title As String, txtname As String, con As String) As String '已有txt但无内容，函数""string返回,已有txt返回第一行内容,写入的con 也为第一行
On Error Resume Next

If Dir("d:\Myuse\" & title & "\" & txtname & ".txt", vbDirectory) = "" Then
Open "d:\Myuse\" & title & "\" & txtname & ".txt" For Output As #1
Print #1, con
Close #1
init_txt = con
Else

Dim TextLine
Dim i As Long
i = 0
Dim a As Integer
a = 1
    Open "d:\Myuse\" & title & "\" & txtname & ".txt" For Input As #1 ' 打开文件。
    
    Do While Not EOF(1) ' 循环至文件尾。
    i = i + 1
    If i = 1 Then
      Line Input #1, TextLine ' 读入一行数据并将其赋予某变量。
      init_txt = TextLine
      a = 0
      Exit Do
    End If

    Loop
    
    Close #1 ' 关闭文件。
                   
If a = 1 Then init_txt = "" '无值返回

End If

End Function
'数组输入为个数一定的数组
Public Function init_txtEx(title As String, txtname As String, ByRef con() As String, num As Integer) As String 'Notice num 不能大于数组的内部个数 ，数组以不一定数[0]开头
'                                          '传人参数时是写数组名，不带括号及其内部的东西
'On Error GoTo Errlog
 init_dir (title)
If Dir("d:\Myuse\" & title & "\" & txtname & ".txt", vbDirectory) = "" Then
Dim k As Integer
Open "d:\Myuse\" & title & "\" & txtname & ".txt" For Append As #1
For k = LBound(con) To (LBound(con) + num - 1)
 'Debug.Print CStr(k)
    Print #1, con(k)
Next k
Close #1

init_txtEx = "w"

Else
Dim TextLine
Dim i As Long
i = LBound(con)
Dim a As Integer
a = 1
    Open "d:\Myuse\" & title & "\" & txtname & ".txt" For Input As #1 ' 打开文件。
    
    Do While Not EOF(1) ' 循环至文件尾。
    
     If (i + 1) <= (LBound(con) + num) Then
      Line Input #1, TextLine ' 读入一行数据并将其赋予某变量。
      con(i) = TextLine
      a = 0
      'Exit Do
     i = i + 1
     Else
      Exit Do
     End If
     
    Loop
    
    Close #1 ' 关闭文件。
                   
If a = 1 Then
init_txtEx = "rd_f" '无值返回
ElseIf a = 0 Then
init_txtEx = "rd"
End If

End If
Exit Function
Errlog:
Call whenErr(Err.number, "AppFilename", "errsection")

End Function



