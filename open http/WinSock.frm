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
   StartUpPosition =   3  '窗口缺省
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
Winsock1.RemoteHost = "202.108.22.5" '设置连接的网址
Winsock1.RemotePort = 80 '设置要连接的远程端口号
Winsock1.Connect '返回与远程计算机的连接。
End Sub

Private Sub Command2_Click()
WebBrowser1.Navigate "www.163.com"

End Sub

Private Sub Winsock1_Connect()

Dim strCommand As String
Dim strWebPage As String
Print "ed"
'当一个 Connect 操作完成时发生
'On Error Resume Next
strWebPage = "http://218.75.21.73" '要下载的文件"
strCommand = "GET " + strWebPage + " HTTP/1.0" + vbCrLf ''GET 为FTP命令 取得文件
strCommand = strCommand + "Accept: */*" + vbCrLf '这句可以不要
'strCommand = strCommand + "Accept: text/html" + vbCrLf '这句可以不要
strCommand = strCommand + vbCrLf '记住一定要加上vbCrLf
Print strCommand
'Debug.Print strCommand '注：你可以用Debug.Print strCommand 来查看一下格式
Winsock1.SendData strCommand '给远程计算机发送数据

End Sub

Private Sub Winsock1_DataArrival(ByVal bytesTotal As Long) '取得数据时产生该事件
t = t + 1

'On Error Resume Next '在错误处理程序结束后，恢复原有的运行
Dim webData As String
Winsock1.GetData webData '检取当前的数据块

Open "D:\5.txt" For Append As #1
Print #1, webData
Close #1
End Sub

Private Sub Form_Load()

End Sub
