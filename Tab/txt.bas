Attribute VB_Name = "txt"
Public Function txtline(address As String, line As Long) As String

Dim TextLine
Dim i As Long
Dim a As Boolean
a = 1
    Open address For Input As #1 ' ���ļ���
    
    Do While Not EOF(1) ' ѭ�����ļ�β��
    i = i + 1
    If i = line Then
      Line Input #1, TextLine ' ����һ�����ݲ����丳��ĳ������
      txtline = TextLine
      a = 0
      Exit Do  '��Ҫ
    End If

    Loop
    
    Close #1 ' �ر��ļ���
                   
If a = 1 Then txtline = "" '��ֵ����

End Function

Public Function txtPrint(title As String, txtname As String, con As String) As Long
Open "d:\Myuse\" & title & "\" & txtname & ".txt" For Output As #1
Print #1, con
Close #1
End Function
