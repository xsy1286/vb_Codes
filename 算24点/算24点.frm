VERSION 5.00
Begin VB.Form Form1 
   BackColor       =   &H80000012&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "24点计算器"
   ClientHeight    =   6735
   ClientLeft      =   45
   ClientTop       =   345
   ClientWidth     =   10215
   Icon            =   "算24点.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   Picture         =   "算24点.frx":0A02
   ScaleHeight     =   6735
   ScaleWidth      =   10215
   StartUpPosition =   3  '窗口缺省
   Begin VB.CommandButton Command2 
      Caption         =   "游戏1"
      Height          =   375
      Left            =   120
      TabIndex        =   15
      Top             =   1080
      Width           =   735
   End
   Begin VB.ListBox List3 
      Height          =   3480
      Left            =   5520
      TabIndex        =   14
      Top             =   1920
      Width           =   2775
   End
   Begin VB.Timer Timer2 
      Interval        =   50
      Left            =   0
      Top             =   2640
   End
   Begin VB.ListBox List2 
      Height          =   3480
      Left            =   3240
      TabIndex        =   13
      Top             =   1920
      Width           =   2295
   End
   Begin VB.OptionButton Option2 
      BackColor       =   &H80000012&
      Caption         =   "其他模式"
      ForeColor       =   &H8000000D&
      Height          =   255
      Left            =   8640
      TabIndex        =   11
      Top             =   2640
      Width           =   1095
   End
   Begin VB.OptionButton Option1 
      BackColor       =   &H80000012&
      Caption         =   "扑克模式"
      ForeColor       =   &H8000000D&
      Height          =   255
      Left            =   8640
      TabIndex        =   10
      Top             =   2160
      Width           =   1095
   End
   Begin VB.TextBox Text6 
      Height          =   375
      Left            =   9000
      TabIndex        =   8
      Text            =   "13"
      Top             =   960
      Width           =   375
   End
   Begin VB.TextBox Text5 
      Height          =   375
      Left            =   9000
      TabIndex        =   7
      Text            =   "1"
      Top             =   360
      Width           =   375
   End
   Begin VB.Timer Timer1 
      Left            =   0
      Top             =   2280
   End
   Begin VB.CommandButton Command1 
      Caption         =   "计算"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   15
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   4200
      TabIndex        =   6
      Top             =   5880
      Width           =   2055
   End
   Begin VB.ListBox List1 
      Height          =   3480
      Left            =   1080
      TabIndex        =   5
      Top             =   1920
      Width           =   2175
   End
   Begin VB.TextBox Text4 
      Height          =   375
      Left            =   6720
      TabIndex        =   3
      Top             =   1080
      Width           =   975
   End
   Begin VB.TextBox Text3 
      Height          =   375
      Left            =   5040
      TabIndex        =   2
      Top             =   1080
      Width           =   975
   End
   Begin VB.TextBox Text2 
      Height          =   375
      Left            =   3360
      TabIndex        =   1
      Top             =   1080
      Width           =   975
   End
   Begin VB.TextBox Text1 
      Height          =   375
      Left            =   1680
      TabIndex        =   0
      Top             =   1080
      Width           =   855
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "           请输入4个数"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   18
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000D&
      Height          =   495
      Left            =   1200
      TabIndex        =   4
      Top             =   240
      Width           =   6615
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "注：范围为正整数"
      ForeColor       =   &H0000FF00&
      Height          =   375
      Left            =   8640
      TabIndex        =   12
      Top             =   1440
      Width           =   1215
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "  范          围"
      ForeColor       =   &H000000FF&
      Height          =   1575
      Left            =   9480
      TabIndex        =   9
      Top             =   120
      Width           =   255
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim x(1 To 4), a, b, c, d, k, i, j, r, bl, g, same(0 To 80), tail As Integer
Dim v As String
Dim s(0 To 80) As String
Dim t As Integer


Private Sub Command1_Click()
x(1) = Val(Text1.Text): x(2) = Val(Text2.Text): x(3) = Val(Text3.Text): x(4) = Val(Text4.Text)
i = Val(Text5.Text): j = Val(Text6.Text)

If i <= j And i > 0 Then

If x(1) <= j And x(1) >= i And x(2) <= j And x(2) >= i And x(3) <= j And x(3) >= i And x(4) <= j And x(4) >= i Then
k = 1
t = 0
Timer1.Interval = 50
Timer1.Enabled = True



'剩下的是交换数据，再复制所有有注释的代码



Else: Label1.Caption = "输入错误，不符合范围，请重新输入"
End If
Else: Label1.Caption = "输入范围错误，请重新输入"
End If
End Sub

Private Sub Command1_KeyPress(KeyAscii As Integer)
KeyAscii = 50
End Sub

Private Sub Command2_Click()
Form2.Show
End Sub

Private Sub Form_Load()


Me.Left = (GetSystemMetrics(SM_CXSCREEN) * Screen.TwipsPerPixelX - Me.Width) / 2 'screen.width必须用此替代
Me.Top = (GetSystemMetrics(SM_CYSCREEN) * Screen.TwipsPerPixelY - Me.Height) / 2


r = 1: g = 1: bl = 1
List3.AddItem "   计算过程（有可能重复）"
'List3.AddItem "！目前！本服务仅供付费用户使用"
List1.AddItem "    输入的数字"
List2.AddItem " 能否算出" + "   " + "基本方法数"
Option1.Value = True
End Sub

Private Sub Option1_Click()
If Option1.Value = True Then Text5.Text = 1: Text6.Text = 13
End Sub

Private Sub Option2_Click()
If Option2.Value = True Then Text5.Text = 1: Text6.Text = 10
End Sub


Private Sub Text1_KeyUp(KeyCode As Integer, Shift As Integer)
If KeyCode = 13 Then Call Command1_Click
End Sub


Private Sub Text2_KeyUp(KeyCode As Integer, Shift As Integer)
If KeyCode = 13 Then Call Command1_Click
End Sub
Private Sub Text3_KeyUp(KeyCode As Integer, Shift As Integer)
If KeyCode = 13 Then Call Command1_Click
End Sub


Private Sub Text4_KeyUp(KeyCode As Integer, Shift As Integer)
If KeyCode = 13 Then Call Command1_Click
End Sub

Private Sub Timer1_Timer()

Label1.Caption = "请稍等片刻"

't = 0
Call col(k, x(1), x(2), x(3), x(4), t, s)

k = k + 1
If k >= 25 Then
            If t = 0 Then v = "否" Else v = "能"
            Timer1.Enabled = False
            tail = 0
            For i = 1 To t
                j = i
                Do While j < t
                    j = j + 1
                    If s(i) = s(j) Then same(i) = 1: j = t: tail = tail + 1  's(i),s(j)是string显示过程内容
                Loop
                If same(i) = 0 Then List3.AddItem (s(i))
            Next
            t = t - tail

Label1.Caption = "        请输入4个数"
List1.AddItem "  " + Str(x(1)) + "   " + Str(x(2)) + "   " + Str(x(3)) + "   " + Str(x(4))
List2.AddItem "    " + v + "         " + Str(t)
End If
End Sub

Private Sub Timer2_Timer()
If z = 0 Then
r = r + 1: If g > 0 Then g = g - 1
End If

If z = 1 Then
g = g + 1: If bl > 0 Then bl = bl - 1
End If
If z = 2 Then bl = bl + 1: r = r - 1

Label1.ForeColor = RGB(r, g, bl)

If r = 250 Then z = 1
If g = 250 Then z = 2
If r = 1 Then z = 0
End Sub
