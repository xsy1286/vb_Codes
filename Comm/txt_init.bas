Attribute VB_Name = "txt_init"
Public Function init_dir(title As String) As Long

   If Dir("D:\Myuse", vbDirectory) = "" Then
   MkDir ("D:\Myuse")
   End If
   
    If Dir("D:\Myuse\" & title, vbDirectory) = "" Then
    MkDir ("D:\Myuse\" & title)
    End If
    
End Function

Public Function init_txt(title As String, txtname As String, con As String) As String
On Error Resume Next

If Dir("d:\Myuse\" & title & "\" & txtname & ".txt", vbDirectory) = "" Then
Open "d:\Myuse\" & title & "\" & txtname & ".txt" For Output As #1
Print #1, con
Close #1
init_txt = con
Else

Dim TextLine
Dim i As Long
Dim a As Integer
a = 1
    Open "d:\Myuse\" & title & "\" & txtname & ".txt" For Input As #1 ' ���ļ���
    
    Do While Not EOF(1) ' ѭ�����ļ�β��
    i = i + 1
    If i = 1 Then
      Line Input #1, TextLine ' ����һ�����ݲ����丳��ĳ������
      init_txt = TextLine
      a = 0
      Exit Do
    End If

    Loop
    
    Close #1 ' �ر��ļ���
                   
If a = 1 Then init_txt = "" '��ֵ����

End If

End Function




