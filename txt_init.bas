Attribute VB_Name = "txt_init"
Option Explicit
Public Function init_dir(title As String) As Long  '�����������ļ���Dir

   If Dir("D:\Myuse", vbDirectory) = "" Then
   MkDir ("D:\Myuse")
   End If
   
    If Dir("D:\Myuse\" & title, vbDirectory) = "" Then
    MkDir ("D:\Myuse\" & title)
    End If
    
End Function

Public Function init_txt(title As String, txtname As String, con As String) As String '����txt�������ݣ�����""string����,����txt���ص�һ������,д���con ҲΪ��һ��
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
'��������Ϊ����һ��������
Public Function init_txtEx(title As String, txtname As String, ByRef con() As String, num As Integer) As String 'Notice num ���ܴ���������ڲ����� �������Բ�һ����[0]��ͷ
'                                          '���˲���ʱ��д���������������ż����ڲ��Ķ���
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
    Open "d:\Myuse\" & title & "\" & txtname & ".txt" For Input As #1 ' ���ļ���
    
    Do While Not EOF(1) ' ѭ�����ļ�β��
    
     If (i + 1) <= (LBound(con) + num) Then
      Line Input #1, TextLine ' ����һ�����ݲ����丳��ĳ������
      con(i) = TextLine
      a = 0
      'Exit Do
     i = i + 1
     Else
      Exit Do
     End If
     
    Loop
    
    Close #1 ' �ر��ļ���
                   
If a = 1 Then
init_txtEx = "rd_f" '��ֵ����
ElseIf a = 0 Then
init_txtEx = "rd"
End If

End If
Exit Function
Errlog:
Call whenErr(Err.number, "AppFilename", "errsection")

End Function



