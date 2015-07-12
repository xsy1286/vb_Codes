VERSION 5.00
Begin VB.Form Form1 
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   3030
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   4560
   LinkTopic       =   "Form1"
   ScaleHeight     =   3030
   ScaleWidth      =   4560
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  '窗口缺省
   Begin VB.Timer Timer1 
      Left            =   3000
      Top             =   360
   End
   Begin VB.TextBox Text1 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   12
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000D&
      Height          =   375
      Index           =   2
      Left            =   600
      TabIndex        =   3
      Text            =   "Text1"
      Top             =   840
      Width           =   615
   End
   Begin VB.TextBox Text1 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   12
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000D&
      Height          =   375
      Index           =   1
      Left            =   360
      TabIndex        =   2
      Text            =   "Text1"
      Top             =   480
      Width           =   1095
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Command1"
      Height          =   615
      Left            =   2280
      TabIndex        =   1
      Top             =   1080
      Visible         =   0   'False
      Width           =   1815
   End
   Begin VB.TextBox Text1 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   12
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000D&
      Height          =   375
      Index           =   0
      Left            =   600
      TabIndex        =   0
      Text            =   "Text1"
      Top             =   120
      Width           =   615
   End
   Begin VB.Label Label1 
      Height          =   135
      Index           =   6
      Left            =   480
      TabIndex        =   10
      Top             =   150
      Width           =   850
   End
   Begin VB.Label Label1 
      Height          =   140
      Index           =   5
      Left            =   495
      TabIndex        =   9
      Top             =   1070
      Width           =   810
   End
   Begin VB.Label Label1 
      Height          =   1250
      Index           =   4
      Left            =   720
      TabIndex        =   8
      Top             =   50
      Width           =   375
   End
   Begin VB.Label Label1 
      Height          =   780
      Index           =   3
      Left            =   1200
      TabIndex        =   7
      Top             =   280
      Width           =   255
   End
   Begin VB.Label Label1 
      Height          =   785
      Index           =   2
      Left            =   360
      TabIndex        =   6
      Top             =   280
      Width           =   255
   End
   Begin VB.Label Label1 
      Height          =   375
      Index           =   1
      Left            =   280
      TabIndex        =   5
      Top             =   480
      Width           =   1250
   End
   Begin VB.Label Label1 
      Height          =   615
      Index           =   0
      Left            =   360
      TabIndex        =   4
      Top             =   360
      Width           =   1095
   End
   Begin VB.Shape Shape1 
      BackColor       =   &H80000007&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H00000000&
      Height          =   1335
      Left            =   0
      Shape           =   3  'Circle
      Top             =   0
      Width           =   1815
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim i As Integer


Private Sub Command1_Click()
transpa Me 'Sub参数不用括号，Function的参数需要
End Sub

Private Sub Form_Load()
transpa Me

init_dir ("Tab")
r = Val(init_txt("Tab", "colorR", "1"))
g = Val(init_txt("Tab", "colorG", "1"))
b = Val(init_txt("Tab", "colorB", "66"))
For i = 0 To 2
Text1(i).Text = ""
Text1(i).ForeColor = RGB(r, g, b)
Next i
For i = 0 To 6
Label1(i).BackColor = vbBlack
Next i
Timer1.Interval = 200

mid_Form Me
End Sub

Private Sub Label1_DblClick(Index As Integer)
Form2.Show  '点到Form2就会Form2显示
End Sub

Private Sub Label1_MouseDown(Index As Integer, Button As Integer, Shift As Integer, x As Single, y As Single)
DoDrag Me
End Sub

Private Sub Text1_DblClick(Index As Integer)
 Form2.Show
End Sub

Private Sub Text1_MouseDown(Index As Integer, Button As Integer, Shift As Integer, x As Single, y As Single)
DoDrag Me
End Sub

Private Sub Timer1_Timer()

betop Me

End Sub
