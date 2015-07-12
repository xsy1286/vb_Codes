Attribute VB_Name = "txt"
Public Function txtPrint(title As String, txtname As String, ByVal con As String) As Long
    Open "d:\Myuse\" & title & "\" & txtname & ".txt" For Output As #1
        Print #1, con
    Close #1
End Function

Public Function txtAppend(title As String, txtname As String, into As String) As Long
        '<EhHeader>
        On Error GoTo txtAppend_Err
        '</EhHeader>

    Open "d:\Myuse\" & title & "\" & txtname & ".txt" For Append As #1
         Print #1, into
     Close #1
        '<EhFooter>
        txtAppend = 1
        Exit Function

txtAppend_Err:
            MsgBox "文件名错误，不能含：\/:*?“<>|"
            txtAppend = 0
        '</EhFooter>
End Function
Public Function txtCrt(title As String, txtname As String) As Long
        '<EhHeader>
        On Error GoTo txtCrt_Err
        '</EhHeader>

100     Open "d:\Myuse\" & title & "\" & txtname & ".txt" For Append As #1
102     Close #1
        '<EhFooter>
        txtCrt = 1
        Exit Function

txtCrt_Err:
 MsgBox "文件名错误，不能含：\/:*?“<>|"
    txtCrt = 0
        '</EhFooter>
End Function

Public Function wr_txtEx(title As String, _
                         txtname As String, _
                         ByRef con() As String, _
                         ByVal num As Integer) As String
                         
    Call txtPrint(title, txtname, con(LBound(con)))

    Dim k As Integer

    For k = (LBound(con) + 1) To (LBound(con) + num - 1)
        Call txtAppend(title, txtname, con(k))
    Next k

End Function
Public Function wrTxt(dir As String, _
                         txtname As String, _
                         ByRef con() As String, _
                         ByVal num As Integer) As String
    Dim k
                         
    Open dir & "\" & txtname & ".txt" For Output As #1
        Print #1, con(LBound(con))
    Close #1
    
 For k = (LBound(con) + 1) To (LBound(con) + num - 1)
    Open dir & "\" & txtname & ".txt" For Append As #1
         Print #1, con(k)
     Close #1
   Next k
        
End Function



Public Function rdTxt(dir As String, _
                         txtname As String, _
                         ByRef con() As String, _
                         ByVal num As Integer) As String
                         
Dim TextLine
Dim i As Long
i = LBound(con)
Dim a As Integer
a = 1

If (LBound(con) + num - 1) > UBound(con) Then
    rd_txtEx = "out of UBound"
    Exit Function
End If

    Open dir & "\" & txtname & ".txt" For Input As #1 ' 打开文件。
    
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
rd_txtEx = "rd_f" '无值返回
ElseIf a = 0 Then
rd_txtEx = "rd"
End If

End Function

Public Function txtlinerd(address As String, line As Long) As String

Dim TextLine
Dim i As Long
Dim a As Boolean
a = 1
    Open address For Input As #1 ' 打开文件。
    
    Do While Not EOF(1) ' 循环至文件尾。
    i = i + 1
    If i = line Then
      Line Input #1, TextLine ' 读入一行数据并将其赋予某变量。
      txtlinerd = TextLine
      a = 0
      Exit Do  '需要
    End If

    Loop
    
    Close #1 ' 关闭文件。
                   
If a = 1 Then txtlinerd = "" '无值返回

End Function
Public Function txtAppendAdr(address As String, into As String) As Long
Open address For Append As #1
Print #1, into
Close #1
End Function

Public Function txtline(address As String) As Long 'how much lines

    Dim i   As Long

    Dim tmp As String

    i = 0
    Open address For Input As #1 ' 打开文件。
    
    Do While Not EOF(1) ' 循环至文件尾。
        i = i + 1
        Line Input #1, tmp
    Loop

    Close #1 ' 关闭文件。

    txtline = i

End Function
