VERSION 5.00
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "MSWINSCK.OCX"
Begin VB.Form Form1 
   Caption         =   "�ͻ���"
   ClientHeight    =   5085
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   6870
   LinkTopic       =   "Form1"
   ScaleHeight     =   5085
   ScaleWidth      =   6870
   StartUpPosition =   3  '����ȱʡ
   Begin VB.TextBox Text4 
      Height          =   375
      Left            =   3600
      TabIndex        =   10
      Top             =   2520
      Width           =   1095
   End
   Begin VB.TextBox Text3 
      Height          =   375
      Left            =   3600
      TabIndex        =   9
      Top             =   2040
      Width           =   2055
   End
   Begin VB.CommandButton Command3 
      Caption         =   "����"
      Height          =   495
      Left            =   720
      TabIndex        =   8
      Top             =   2160
      Width           =   1455
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Command2"
      Height          =   1215
      Left            =   4560
      TabIndex        =   3
      Top             =   3120
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.TextBox Text2 
      Height          =   1335
      Left            =   600
      MultiLine       =   -1  'True
      ScrollBars      =   3  'Both
      TabIndex        =   2
      Top             =   3000
      Width           =   3255
   End
   Begin VB.TextBox Text1 
      Height          =   1335
      Left            =   600
      MultiLine       =   -1  'True
      ScrollBars      =   3  'Both
      TabIndex        =   1
      Top             =   600
      Width           =   3375
   End
   Begin MSWinsockLib.Winsock Winsock1 
      Left            =   360
      Top             =   3960
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
   Begin VB.Timer Timer1 
      Left            =   5400
      Top             =   1920
   End
   Begin VB.CommandButton Command1 
      Enabled         =   0   'False
      Height          =   375
      Left            =   5160
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   720
      Width           =   1335
   End
   Begin VB.Label Label7 
      Caption         =   "Label7"
      Height          =   255
      Left            =   3120
      TabIndex        =   13
      Top             =   120
      Width           =   2775
   End
   Begin VB.Label Label6 
      Caption         =   "   �˿ڣ�"
      Height          =   255
      Left            =   2760
      TabIndex        =   12
      Top             =   2520
      Width           =   1095
   End
   Begin VB.Label Label5 
      Caption         =   "IP��ַ��"
      Height          =   375
      Left            =   2880
      TabIndex        =   11
      Top             =   2040
      Width           =   1095
   End
   Begin VB.Label Label4 
      Caption         =   "Label4"
      Height          =   375
      Left            =   1800
      TabIndex        =   7
      Top             =   120
      Width           =   2175
   End
   Begin VB.Label Label3 
      Height          =   495
      Left            =   5760
      TabIndex        =   6
      Top             =   1080
      Width           =   1575
   End
   Begin VB.Label Label2 
      Caption         =   "��  ��  ��"
      Height          =   1095
      Left            =   120
      TabIndex        =   5
      Top             =   3000
      Width           =   255
   End
   Begin VB.Label Label1 
      Caption         =   "��  ʾ  ��"
      Height          =   975
      Left            =   120
      TabIndex        =   4
      Top             =   720
      Width           =   255
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim tmp As String

Private Sub Command2_Click()
Text1.Text = Text1.Text & "1"
End Sub

Private Sub Command3_Click()
Winsock1.Close

Winsock1.RemotePort = Val(Text4.Text)
Winsock1.RemoteHost = Text3.Text

Winsock1.Connect

Label7.Caption = ""
End Sub

Private Sub Form_Load()
Text1.Width = 5000
Text2.Width = 5000
Text3.Text = "123.159.176.220"
Text4.Text = "2066"
Label7.Caption = ""


mid_Form Me

Command1.BackColor = vbRed
Timer1.Interval = 10
'Winsock1.Protocol = sckUDPProtocol


End Sub

Private Sub Text2_Change()

If (Winsock1.State = 7) Then

If (Text2.Text = "") Then
Winsock1.SendData ("zpq,.'[]85zzu8ua")
Else
Winsock1.SendData (Text2.Text)
End If

End If
End Sub



Private Sub Timer1_Timer()

If (Winsock1.State = 7) Then
Command1.BackColor = vbGreen
Else
Command1.BackColor = vbRed
End If

Label3.Caption = "state: " + CStr(Winsock1.State)

If (Winsock1.State = 0) Then
Label4.Caption = "  �ر�"
ElseIf (Winsock1.State = 1) Then
Label4.Caption = "  ��"
ElseIf (Winsock1.State = 3) Then
Label4.Caption = "���ӹ���"
ElseIf (Winsock1.State = 4) Then
Label4.Caption = "ʶ������"
ElseIf (Winsock1.State = 5) Then
Label4.Caption = "��ʶ������"
ElseIf (Winsock1.State = 6) Then
Label4.Caption = "��������"
ElseIf (Winsock1.State = 7) Then
Label4.Caption = " ������"
ElseIf (Winsock1.State = 9) Then
Label4.Caption = "  ����"
Else
Label4.Caption = ""
End If

If (Text2.Text = "") And (Winsock1.State = 7) Then Winsock1.SendData ("")
End Sub


Private Sub Winsock1_DataArrival(ByVal bytesTotal As Long)

If (Winsock1.State = 7) Then
Winsock1.GetData tmp

If (tmp = "zpq,.'[]85zzu8ua") Then
Text1.Text = ""
Else
Text1.Text = tmp
End If

End If

End Sub

Private Sub Winsock1_Error(ByVal Number As Integer, Description As String, ByVal Scode As Long, ByVal Source As String, ByVal HelpFile As String, ByVal HelpContext As Long, CancelDisplay As Boolean)
     Select Case Number   ' �ж� Number ��ֵ��
         Case 7
           Label7.Caption = "�ڴ治��,���Ժ�!"
         
         Case 380, 394, 383
           Label7.Caption = "����ֵ��Ч!"
         Case 1004
           Label7.Caption = "ȡ������!"
         Case 10014
           Label7.Caption = "������ĵ�ַ�ǹ㲥��ַ����δ���ñ��!"
         Case 10035
           Label7.Caption = "�׽��ֲ��ɿ飬��ָ��������ʹ֮�ɿ�!"
         Case 10036
           Label7.Caption = "������ Winsock �����ڽ���֮��!"
         Case 10037
           Label7.Caption = "��ɲ�����δ���������Ĳ���!"
         Case 10038
           Label7.Caption = "�����������׽���!"
         Case 10040
           Label7.Caption = "���ݱ�̫�󣬲����ڻ�������Ҫ��������ض�!"
         Case 10043
           Label7.Caption = "��֧��ָ���Ķ˿�!"
         Case 10048
          Label7.Caption = "��ַ��ʹ����!"
         Case 10049
           Label7.Caption = "���Ա��ػ����Ĳ����õ�ַ!"
         Case 10050
           Label7.Caption = "������ϵͳʧ��!"
         Case 10051
           Label7.Caption = "��ʱ���ܴ�������������!"
         Case 10052
           Label7.Caption = "������ SO_KEEPALIVE ʱ���ӳ�ʱ!"
         Case 11053
          Label7.Caption = "���ڳ�ʱ��������ʧ�ܶ���ֹ����!"
         Case 10054
          Label7.Caption = "ͨ��Զ��������������!"
         Case 10055
           Label7.Caption = "û�п��õĻ���ռ�!"
         Case 10056
           Label7.Caption = "�������׽���!"
         Case 10057
           Label7.Caption = "δ�����׽���!"
         Case 10058
           Label7.Caption = "�ѹر��׽���!"
         Case 10060
           Label7.Caption = "�ѹر��׽���!"
         Case 10061
          Label7.Caption = "ǿ�оܾ�����!"
         Case 10093
           Label7.Caption = "Ӧ���ȵ��� WinsockInit!"
         Case 11001
          Label7.Caption = "��ȨӦ��δ�ҵ�����!"
           
         Case 11002
           Label7.Caption = "����ȨӦ��δ�ҵ�����!"
         Case 11003
          Label7.Caption = "���ɻָ��Ĵ���!"
         Case 11004
           Label7.Caption = "��Ч����������������������ݼ�¼!"
         
         Case Else   ' ������ֵ��
           Debug.Print "Not between 1 and 10"
         End Select
 End Sub

