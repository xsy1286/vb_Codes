VERSION 5.00
Begin VB.Form Form2 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "24点游戏1"
   ClientHeight    =   3645
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   5535
   Icon            =   "Form2.frx":0000
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   Picture         =   "Form2.frx":1CCA
   ScaleHeight     =   3645
   ScaleWidth      =   5535
   StartUpPosition =   3  '窗口缺省
   Begin VB.Timer Timer2 
      Left            =   1680
      Top             =   2280
   End
   Begin VB.Timer Timer1 
      Left            =   1080
      Top             =   2520
   End
   Begin VB.Timer mm 
      Left            =   360
      Top             =   2640
   End
   Begin VB.TextBox Text2 
      Height          =   615
      Left            =   2760
      MultiLine       =   -1  'True
      TabIndex        =   19
      Text            =   "Form2.frx":40ED9
      Top             =   1440
      Width           =   855
   End
   Begin VB.TextBox Text1 
      Height          =   375
      Index           =   16
      Left            =   0
      TabIndex        =   18
      Text            =   "Text1"
      Top             =   0
      Width           =   615
   End
   Begin VB.ListBox List1 
      Height          =   2220
      Left            =   3600
      TabIndex        =   17
      Top             =   360
      Width           =   1695
   End
   Begin VB.CommandButton Command1 
      Caption         =   "计算"
      Height          =   495
      Left            =   1680
      TabIndex        =   16
      Top             =   2880
      Width           =   2295
   End
   Begin VB.TextBox Text1 
      Height          =   375
      Index           =   15
      Left            =   1920
      TabIndex        =   15
      Text            =   "Text1"
      Top             =   840
      Width           =   615
   End
   Begin VB.TextBox Text1 
      Height          =   375
      Index           =   14
      Left            =   0
      TabIndex        =   14
      Text            =   "Text1"
      Top             =   0
      Width           =   615
   End
   Begin VB.TextBox Text1 
      Height          =   375
      Index           =   13
      Left            =   0
      TabIndex        =   13
      Text            =   "Text1"
      Top             =   0
      Width           =   615
   End
   Begin VB.TextBox Text1 
      Height          =   375
      Index           =   12
      Left            =   0
      TabIndex        =   12
      Text            =   "Text1"
      Top             =   0
      Width           =   615
   End
   Begin VB.TextBox Text1 
      Height          =   375
      Index           =   11
      Left            =   0
      TabIndex        =   11
      Text            =   "Text1"
      Top             =   0
      Width           =   615
   End
   Begin VB.TextBox Text1 
      Height          =   375
      Index           =   10
      Left            =   0
      TabIndex        =   10
      Text            =   "Text1"
      Top             =   0
      Width           =   615
   End
   Begin VB.TextBox Text1 
      Height          =   375
      Index           =   9
      Left            =   0
      TabIndex        =   9
      Text            =   "Text1"
      Top             =   0
      Width           =   615
   End
   Begin VB.TextBox Text1 
      Height          =   375
      Index           =   8
      Left            =   0
      TabIndex        =   8
      Text            =   "Text1"
      Top             =   0
      Width           =   615
   End
   Begin VB.TextBox Text1 
      Height          =   375
      Index           =   7
      Left            =   0
      TabIndex        =   7
      Text            =   "Text1"
      Top             =   0
      Width           =   615
   End
   Begin VB.TextBox Text1 
      Height          =   375
      Index           =   6
      Left            =   0
      TabIndex        =   6
      Text            =   "Text1"
      Top             =   0
      Width           =   615
   End
   Begin VB.TextBox Text1 
      Height          =   375
      Index           =   5
      Left            =   0
      TabIndex        =   5
      Text            =   "Text1"
      Top             =   0
      Width           =   615
   End
   Begin VB.TextBox Text1 
      Height          =   375
      Index           =   4
      Left            =   0
      TabIndex        =   4
      Text            =   "Text1"
      Top             =   0
      Width           =   615
   End
   Begin VB.TextBox Text1 
      Height          =   375
      Index           =   3
      Left            =   0
      TabIndex        =   3
      Text            =   "Text1"
      Top             =   0
      Width           =   615
   End
   Begin VB.TextBox Text1 
      Height          =   375
      Index           =   2
      Left            =   0
      TabIndex        =   2
      Text            =   "Text1"
      Top             =   0
      Width           =   615
   End
   Begin VB.TextBox Text1 
      Height          =   375
      Index           =   1
      Left            =   0
      TabIndex        =   1
      Text            =   "Text1"
      Top             =   0
      Width           =   615
   End
   Begin VB.TextBox Text1 
      Height          =   375
      Index           =   0
      Left            =   720
      TabIndex        =   0
      Text            =   "Text1"
      Top             =   840
      Width           =   615
   End
End
Attribute VB_Name = "Form2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim mmt As Double
Dim srt As Boolean
Dim x(1 To 16) As Integer
Dim t1 As Integer
Dim t2 As Integer
Private Const linetime = 900  '依然是时间问题
Private Const singletime = 30  '调试单步时计算量非正常运行量
Dim y(1 To 4) As Integer
Dim brk As Integer
Dim sg As String
Dim game As Integer
Dim tempstr() As String


Private Sub Command1_Click()

Dim i As Integer

For i = 1 To 16
x(i) = Val(Text1(i).Text)
If x(i) = 0 Then MsgBox "不能出现0": Exit Sub
Next i

t1 = 0
Timer1.Interval = linetime
Timer1.Enabled = True

List1.Clear
List1.Visible = True
Text2.Visible = False


game = 0
srt = True

Command1.Caption = "请稍后"
End Sub

Private Sub Form_Load()

Me.Left = (GetSystemMetrics(SM_CXSCREEN) * Screen.TwipsPerPixelX - Me.Width) / 2 'screen.width必须用此替代
Me.Top = (GetSystemMetrics(SM_CYSCREEN) * Screen.TwipsPerPixelY - Me.Height) / 2
List1.Visible = False
Text2.Left = List1.Left
Text2.Width = List1.Width
Text2.Top = List1.Top
Text2.Height = List1.Height
Text1(0).Visible = False
Dim i As Integer
For i = 1 To 16
Text1(i).Width = 600: Text1(i).Height = 600
Text1(i).Left = (i - 4 * ((i - 1) \ 4)) * 600
Text1(i).Top = ((i - 1) \ 4) * 600 + 300
Next i

Text1(1).Text = GetSetting("24", "Data", "1", "")
Text1(2).Text = GetSetting("24", "Data", "2", "")
Text1(3).Text = GetSetting("24", "Data", "3", "")
Text1(4).Text = GetSetting("24", "Data", "4", "")
Text1(5).Text = GetSetting("24", "Data", "5", "")
Text1(6).Text = GetSetting("24", "Data", "6", "")
Text1(7).Text = GetSetting("24", "Data", "7", "")
Text1(8).Text = GetSetting("24", "Data", "8", "")
Text1(9).Text = GetSetting("24", "Data", "9", "")
Text1(10).Text = GetSetting("24", "Data", "10", "")
Text1(11).Text = GetSetting("24", "Data", "11", "")
Text1(12).Text = GetSetting("24", "Data", "12", "")
Text1(13).Text = GetSetting("24", "Data", "13", "")
Text1(14).Text = GetSetting("24", "Data", "14", "")
Text1(15).Text = GetSetting("24", "Data", "15", "")
Text1(16).Text = GetSetting("24", "Data", "16", "")

sg = ""

srt = False
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
If srt = True Then
If mm.Enabled <> False Then
mm.Interval = 0
mmt = 0
mm.Enabled = False
List1.Visible = 1
Text2.Visible = 0
End If
End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
Call SaveSetting("24", "Data", "1", Text1(1).Text)
Call SaveSetting("24", "Data", "2", Text1(2).Text)
Call SaveSetting("24", "Data", "3", Text1(3).Text)
Call SaveSetting("24", "Data", "4", Text1(4).Text)
Call SaveSetting("24", "Data", "5", Text1(5).Text)
Call SaveSetting("24", "Data", "6", Text1(6).Text)
Call SaveSetting("24", "Data", "7", Text1(7).Text)
Call SaveSetting("24", "Data", "8", Text1(8).Text)
Call SaveSetting("24", "Data", "9", Text1(9).Text)
Call SaveSetting("24", "Data", "10", Text1(10).Text)
Call SaveSetting("24", "Data", "11", Text1(11).Text)
Call SaveSetting("24", "Data", "12", Text1(12).Text)
Call SaveSetting("24", "Data", "13", Text1(13).Text)
Call SaveSetting("24", "Data", "14", Text1(14).Text)
Call SaveSetting("24", "Data", "15", Text1(15).Text)
Call SaveSetting("24", "Data", "16", Text1(16).Text)

End Sub

Private Sub List1_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
mm.Interval = 50
mm.Enabled = True

End Sub

Private Sub mm_Timer()

If mmt > 60 Then
List1.Visible = False
Text2.Visible = True
Else
mmt = mmt + 1
End If

End Sub


Private Sub Timer1_Timer()
t1 = t1 + 1
t2 = 0
brk = 0
Select Case t1
Case 1
           sg = "第一行错误"
           y(1) = x(1): y(2) = x(2): y(3) = x(3): y(4) = x(4): Timer2.Interval = singletime: Timer2.Enabled = True
Case 2
           sg = "第二行错误"
           y(1) = x(5): y(2) = x(6): y(3) = x(7): y(4) = x(8): Timer2.Interval = singletime: Timer2.Enabled = True
Case 3
           sg = "第三行错误"
           y(1) = x(9): y(2) = x(10): y(3) = x(11): y(4) = x(12): Timer2.Interval = singletime: Timer2.Enabled = True
Case 4
           sg = "第四行错误"
           y(1) = x(13): y(2) = x(14): y(3) = x(15): y(4) = x(16): Timer2.Interval = singletime: Timer2.Enabled = True
Case 5
           sg = "第一列错误"
           y(1) = x(1): y(2) = x(5): y(3) = x(9): y(4) = x(13): Timer2.Interval = singletime: Timer2.Enabled = True
Case 6
           sg = "第二列错误"
           y(1) = x(2): y(2) = x(6): y(3) = x(10): y(4) = x(14): Timer2.Interval = singletime: Timer2.Enabled = True
Case 7
           sg = "第三列错误"
           y(1) = x(3): y(2) = x(7): y(3) = x(11): y(4) = x(15): Timer2.Interval = singletime: Timer2.Enabled = True
Case 8
           sg = "第四列错误"
           y(1) = x(4): y(2) = x(2): y(3) = x(12): y(4) = x(16): Timer2.Interval = singletime: Timer2.Enabled = True
Case 9
           sg = "捺对角错误"
           y(1) = x(1): y(2) = x(6): y(3) = x(11): y(4) = x(16): Timer2.Interval = singletime: Timer2.Enabled = True
Case 10
           sg = "撇对角错误"
           y(1) = x(4): y(2) = x(7): y(3) = x(10): y(4) = x(13): Timer2.Interval = singletime: Timer2.Enabled = True
End Select



If t1 >= 11 Then

If game = 0 Then
List1.Clear
List1.AddItem "You Win"
Else
List1.AddItem "You Lose"
End If
t1 = 0
Timer1.Enabled = False
Command1.Caption = "计算"
'Timer1.Interval = 0:
End If


End Sub

Private Sub Timer2_Timer()
t2 = t2 + 1

Debug.Print "t2:" & CStr(t2)

If t2 >= 25 Then

If brk > 0 Then
List1.AddItem sg
game = game + 1
End If


Timer2.Enabled = False: t2 = 0

Else

Call col(t2, y(1), y(2), y(3), y(4), brk, tempstr())
End If

End Sub
