VERSION 5.00
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "MSWINSCK.OCX"
Begin VB.Form Form1 
   Caption         =   "客户端"
   ClientHeight    =   5085
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   6870
   LinkTopic       =   "Form1"
   ScaleHeight     =   5085
   ScaleWidth      =   6870
   StartUpPosition =   3  '窗口缺省
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
      Caption         =   "连接"
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
      Caption         =   "   端口："
      Height          =   255
      Left            =   2760
      TabIndex        =   12
      Top             =   2520
      Width           =   1095
   End
   Begin VB.Label Label5 
      Caption         =   "IP地址："
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
      Caption         =   "输  入  框"
      Height          =   1095
      Left            =   120
      TabIndex        =   5
      Top             =   3000
      Width           =   255
   End
   Begin VB.Label Label1 
      Caption         =   "显  示  框"
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
Label4.Caption = "  关闭"
ElseIf (Winsock1.State = 1) Then
Label4.Caption = "  打开"
ElseIf (Winsock1.State = 3) Then
Label4.Caption = "连接挂起"
ElseIf (Winsock1.State = 4) Then
Label4.Caption = "识别主机"
ElseIf (Winsock1.State = 5) Then
Label4.Caption = "已识别主机"
ElseIf (Winsock1.State = 6) Then
Label4.Caption = "正在连接"
ElseIf (Winsock1.State = 7) Then
Label4.Caption = " 已连接"
ElseIf (Winsock1.State = 9) Then
Label4.Caption = "  错误"
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
     Select Case Number   ' 判断 Number 的值。
         Case 7
           Label7.Caption = "内存不足,请稍候!"
         
         Case 380, 394, 383
           Label7.Caption = "属性值无效!"
         Case 1004
           Label7.Caption = "取消操作!"
         Case 10014
           Label7.Caption = "所请求的地址是广播地址，但未设置标记!"
         Case 10035
           Label7.Caption = "套接字不成块，而指定操作将使之成块!"
         Case 10036
           Label7.Caption = "制造块的 Winsock 操作在进行之中!"
         Case 10037
           Label7.Caption = "完成操作。未进行制造块的操作!"
         Case 10038
           Label7.Caption = "描述符不是套接字!"
         Case 10040
           Label7.Caption = "数据报太大，不适于缓冲区的要求，因而被截断!"
         Case 10043
           Label7.Caption = "不支持指定的端口!"
         Case 10048
          Label7.Caption = "地址在使用中!"
         Case 10049
           Label7.Caption = "来自本地机器的不可用地址!"
         Case 10050
           Label7.Caption = "网络子系统失败!"
         Case 10051
           Label7.Caption = "此时不能从主机到达网络!"
         Case 10052
           Label7.Caption = "在设置 SO_KEEPALIVE 时连接超时!"
         Case 11053
          Label7.Caption = "由于超时或者其它失败而中止连接!"
         Case 10054
          Label7.Caption = "通过远端重新设置连接!"
         Case 10055
           Label7.Caption = "没有可用的缓冲空间!"
         Case 10056
           Label7.Caption = "已连接套接字!"
         Case 10057
           Label7.Caption = "未连接套接字!"
         Case 10058
           Label7.Caption = "已关闭套接字!"
         Case 10060
           Label7.Caption = "已关闭套接字!"
         Case 10061
          Label7.Caption = "强行拒绝连接!"
         Case 10093
           Label7.Caption = "应首先调用 WinsockInit!"
         Case 11001
          Label7.Caption = "授权应答：未找到主机!"
           
         Case 11002
           Label7.Caption = "非授权应答：未找到主机!"
         Case 11003
          Label7.Caption = "不可恢复的错误!"
         Case 11004
           Label7.Caption = "无效名，对所请求的类型无数据记录!"
         
         Case Else   ' 其他数值。
           Debug.Print "Not between 1 and 10"
         End Select
 End Sub

