VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   7695
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   9270
   LinkTopic       =   "Form1"
   ScaleHeight     =   7695
   ScaleWidth      =   9270
   StartUpPosition =   3  '����ȱʡ
   Begin VB.PictureBox Winsock1 
      Height          =   480
      Left            =   840
      ScaleHeight     =   420
      ScaleWidth      =   1140
      TabIndex        =   5
      Top             =   3000
      Width           =   1200
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Command2"
      Height          =   615
      Left            =   5880
      TabIndex        =   4
      Top             =   4680
      Width           =   2415
   End
   Begin VB.PictureBox WebBrowser1 
      Height          =   2535
      Left            =   2160
      ScaleHeight     =   2475
      ScaleWidth      =   5115
      TabIndex        =   3
      Top             =   3840
      Width           =   5175
   End
   Begin VB.TextBox Text2 
      Height          =   615
      Left            =   7200
      TabIndex        =   2
      Text            =   "Text2"
      Top             =   3600
      Width           =   1575
   End
   Begin VB.TextBox Text1 
      Height          =   1935
      Left            =   2280
      MultiLine       =   -1  'True
      TabIndex        =   1
      Text            =   "WinSock.frx":0000
      Top             =   840
      Width           =   5055
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Command1"
      Height          =   615
      Left            =   600
      TabIndex        =   0
      Top             =   4200
      Width           =   2415
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim t As Integer
Private Sub Command1_Click()
Winsock1.RemoteHost = "202.108.22.5" '�������ӵ���ַ
Winsock1.RemotePort = 80 '����Ҫ���ӵ�Զ�̶˿ں�
Winsock1.Connect '������Զ�̼���������ӡ�
End Sub

Private Sub Command2_Click()
WebBrowser1.Navigate "www.163.com"

End Sub

Private Sub Winsock1_Connect()

Dim strCommand As String
Dim strWebPage As String
Print "ed"
'��һ�� Connect �������ʱ����
'On Error Resume Next
strWebPage = "http://218.75.21.73" 'Ҫ���ص��ļ�"
strCommand = "GET " + strWebPage + " HTTP/1.0" + vbCrLf ''GET ΪFTP���� ȡ���ļ�
strCommand = strCommand + "Accept: */*" + vbCrLf '�����Բ�Ҫ
'strCommand = strCommand + "Accept: text/html" + vbCrLf '�����Բ�Ҫ
strCommand = strCommand + vbCrLf '��סһ��Ҫ����vbCrLf
Print strCommand
'Debug.Print strCommand 'ע���������Debug.Print strCommand ���鿴һ�¸�ʽ
Winsock1.SendData strCommand '��Զ�̼������������

End Sub

Private Sub Winsock1_DataArrival(ByVal bytesTotal As Long) 'ȡ������ʱ�������¼�
t = t + 1

'On Error Resume Next '�ڴ������������󣬻ָ�ԭ�е�����
Dim webData As String
Winsock1.GetData webData '��ȡ��ǰ�����ݿ�

Open "D:\5.txt" For Append As #1
Print #1, webData
Close #1
End Sub

Private Sub Form_Load()

End Sub
