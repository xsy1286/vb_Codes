Attribute VB_Name = "txt"
Public Function txtline(address As String, line As Long) As String

Dim TextLine
Dim i As Long
Dim a As Boolean
a = 1
    Open address For Input As #1 ' 打开文件。
    
    Do While Not EOF(1) ' 循环至文件尾。
    i = i + 1
    If i = line Then
      Line Input #1, TextLine ' 读入一行数据并将其赋予某变量。
      txtline = TextLine
      a = 0
      Exit Do  '需要
    End If

    Loop
    
    Close #1 ' 关闭文件。
                   
If a = 1 Then txtline = "" '无值返回

End Function

Public Function txtPrint(title As String, txtname As String, con As String) As Long
Open "d:\Myuse\" & title & "\" & txtname & ".txt" For Output As #1
Print #1, con
Close #1
End Function
