VERSION 5.00
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "MSWINSCK.OCX"
Begin VB.Form Form1 
   Caption         =   "服务端"
   ClientHeight    =   4965
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   6690
   Icon            =   "Form1.frx":0000
   LinkTopic       =   "Form1"
   Picture         =   "Form1.frx":014A
   ScaleHeight     =   4965
   ScaleWidth      =   6690
   StartUpPosition =   3  '窗口缺省
   Begin VB.CommandButton Command1 
      Caption         =   "重设端口"
      Height          =   495
      Left            =   5640
      TabIndex        =   5
      Top             =   3480
      Width           =   975
   End
   Begin VB.TextBox Text3 
      Height          =   375
      Left            =   5640
      TabIndex        =   4
      Text            =   "Text3"
      Top             =   3000
      Width           =   975
   End
   Begin VB.TextBox Text2 
      Height          =   1455
      Left            =   600
      MultiLine       =   -1  'True
      ScrollBars      =   3  'Both
      TabIndex        =   3
      Top             =   3240
      Width           =   3135
   End
   Begin VB.Timer Timer1 
      Left            =   4200
      Top             =   2280
   End
   Begin VB.CommandButton Command2 
      Caption         =   "本机IP"
      Height          =   1455
      Left            =   360
      TabIndex        =   1
      Top             =   3240
      Width           =   255
   End
   Begin MSWinsockLib.Winsock Winsock1 
      Left            =   3600
      Top             =   2280
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
   Begin VB.TextBox Text1 
      Height          =   1575
      Left            =   480
      MultiLine       =   -1  'True
      ScrollBars      =   3  'Both
      TabIndex        =   0
      Top             =   240
      Width           =   3975
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Label1"
      Height          =   255
      Left            =   4320
      TabIndex        =   2
      Top             =   2280
      Width           =   975
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim tmp As String

Private Sub Command1_Click()
Winsock1.Close
Winsock1.LocalPort = Val(Text3.Text)
Winsock1.Listen
End Sub

'Dim getid As Long
Private Sub Command2_Click()
Print Winsock1.LocalIP
End Sub

Private Sub Form_Load()

mid_Form Me

Text1.Width = 5000
Text2.Width = 5000
'Winsock1.Protocol = sckUDPProtocol
Winsock1.LocalPort = 8999
Winsock1.Listen
Text3.Text = CStr(Winsock1.LocalPort)
Timer1.Interval = 10
End Sub

Private Sub Text2_Change()
Debug.Print "chage"
If (Winsock1.State = 7) Then
Debug.Print "send"
If (Text2.Text = "") Then
Winsock1.SendData ("zpq,.'[]85zzu8ua")
Else
Winsock1.SendData (Text2.Text)
End If

End If
End Sub

Private Sub Timer1_Timer()
Label1.Caption = "state: " + CStr(Winsock1.State)

If (Winsock1.State = 8) Then
Winsock1.Close
Winsock1.Listen
End If

End Sub

Private Sub Winsock1_ConnectionRequest(ByVal requestID As Long)
   If Winsock1.State <> sckClosed Then Winsock1.Close

   Winsock1.Accept requestID
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
