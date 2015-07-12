VERSION 5.00
Begin VB.Form Form1 
   AutoRedraw      =   -1  'True
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   3720
   ClientLeft      =   0
   ClientTop       =   996
   ClientWidth     =   4032
   Icon            =   "Form1.frx":0000
   LinkTopic       =   "Form1"
   Picture         =   "Form1.frx":1CCA
   ScaleHeight     =   3720
   ScaleWidth      =   4032
   ShowInTaskbar   =   0   'False
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   252
      Left            =   1560
      Picture         =   "Form1.frx":6228C
      ScaleHeight     =   252
      ScaleWidth      =   600
      TabIndex        =   14
      Top             =   360
      Width           =   600
   End
   Begin VB.TextBox Text1 
      Height          =   264
      Left            =   360
      TabIndex        =   15
      Top             =   1440
      Width           =   2052
   End
   Begin VB.Timer Timer3 
      Left            =   720
      Top             =   3360
   End
   Begin VB.Timer Timer2 
      Left            =   360
      Top             =   3360
   End
   Begin VB.CommandButton Command2 
      BackColor       =   &H00FFFF80&
      Caption         =   "确定"
      Height          =   372
      Left            =   600
      Style           =   1  'Graphical
      TabIndex        =   10
      Top             =   3000
      Width           =   852
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H00FFFF80&
      Caption         =   "取消"
      Height          =   372
      Left            =   2280
      Style           =   1  'Graphical
      TabIndex        =   9
      Top             =   3000
      Width           =   852
   End
   Begin VB.ComboBox Combo3 
      Height          =   276
      Left            =   2760
      TabIndex        =   8
      Text            =   "Combo3"
      Top             =   2160
      Width           =   732
   End
   Begin VB.ComboBox Combo2 
      Height          =   276
      Left            =   1560
      TabIndex        =   7
      Text            =   "Combo2"
      Top             =   2160
      Width           =   732
   End
   Begin VB.ComboBox Combo1 
      Height          =   276
      Left            =   240
      TabIndex        =   6
      Text            =   "Combo1"
      Top             =   2160
      Width           =   972
   End
   Begin VB.CheckBox Check1 
      BackColor       =   &H80000012&
      Caption         =   "移动界面"
      ForeColor       =   &H00C0C0C0&
      Height          =   252
      Left            =   2640
      TabIndex        =   0
      Top             =   1440
      Width           =   1092
   End
   Begin VB.Timer Timer1 
      Left            =   0
      Top             =   3360
   End
   Begin VB.Label Label8 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "日"
      ForeColor       =   &H00C0C0C0&
      Height          =   180
      Left            =   3600
      TabIndex        =   13
      Top             =   2160
      Width           =   180
   End
   Begin VB.Label Label7 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "月"
      ForeColor       =   &H00C0C0C0&
      Height          =   180
      Left            =   2400
      TabIndex        =   12
      Top             =   2160
      Width           =   180
   End
   Begin VB.Label Label6 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "年"
      ForeColor       =   &H00C0C0C0&
      Height          =   180
      Left            =   1320
      TabIndex        =   11
      Top             =   2160
      Width           =   180
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "天"
      ForeColor       =   &H00C0C0C0&
      Height          =   180
      Left            =   3240
      TabIndex        =   5
      Top             =   120
      Width           =   180
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Label4"
      BeginProperty Font 
         Name            =   "隶书"
         Size            =   12
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   240
      Left            =   2280
      TabIndex        =   4
      Top             =   120
      Width           =   720
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "还有"
      ForeColor       =   &H00C0C0C0&
      Height          =   180
      Left            =   1680
      TabIndex        =   3
      Top             =   120
      Width           =   360
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackColor       =   &H8000000D&
      BackStyle       =   0  'Transparent
      Caption         =   "Label2"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   9
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000D&
      Height          =   180
      Left            =   720
      TabIndex        =   2
      Top             =   120
      Width           =   648
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "距"
      ForeColor       =   &H00C0C0C0&
      Height          =   180
      Left            =   240
      TabIndex        =   1
      Top             =   120
      Width           =   180
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Type dat
y As String
m As String
d As String
End Type
Dim i As Long
Private Type POINTAPI
    X As Long
    y As Long
End Type
Private Declare Function GetCursorPos Lib "user32" (lpoint As POINTAPI) As Long

Dim pa As POINTAPI '！！！定义结构必须分开定义
Dim pv As POINTAPI
Dim tp As POINTAPI
Dim tp2 As POINTAPI
Private Declare Function SetWindowPos Lib "user32" (ByVal hwnd As Long, ByVal hWndInsertAfter As Long, ByVal X As Long, ByVal y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long) As Long
Const HWND_TOPMOST = -1
Const HWND_NOTOPMOST = -2
Const SWP_NOMOVE = &H2
Const SWP_NOSIZE = &H1
Const SWP_SHOWWINDOW = &H40
 Const SWP_NOOWNERZORDER = &H200      '  Don't do owner Z ordering
 Const SWP_NOACTIVATE = &H10
 Dim tmp As String
Dim top1 As Boolean
Dim d As Integer
Dim t As Long
Dim t2 As Integer
Dim bt2 As Integer
Dim dh As Integer
Dim nock As Integer
Dim date1 As dat
Dim cv As Integer



Private Sub Combo3_DropDown()
'combo3.Clear
Debug.Print "3"

Dim y As Integer
Dim m As Integer
m = Val(Combo2.Text)
y = Val(Combo1.Text)
For i = 1 To 9
Combo3.AddItem "0" & Mid(Str(i), 2)
Next i

If m = 2 Then

If run(y) = True Then
For i = 10 To 29
Combo3.AddItem Mid(Str(i), 2)
Next i
Else
For i = 10 To 28
Combo3.AddItem Mid(Str(i), 2)
Next i
End If

ElseIf m = 4 Or m = 6 Or m = 9 Or m = 11 Then
For i = 10 To 30
Combo3.AddItem Mid(Str(i), 2)
Next i

Else
For i = 10 To 31
Combo3.AddItem Mid(Str(i), 2)
Next i

End If

End Sub

Private Sub Command1_Click()
If bt2 = 2 Then
bt2 = 3
Timer3.Interval = 20
Timer3.Enabled = True
End If
Combo1.Text = date1.y
Combo2.Text = date1.m
Combo3.Text = date1.d
Check1.Value = cv
Text1.Text = s

End Sub

Private Sub Command2_Click()
Dim y As Integer
Dim m As Integer
Dim d As Integer
y = Val(Combo1.Text)
m = Val(Combo2.Text)
d = Val(Combo3.Text)

Dim doit  As Boolean
doit = True

If m = 2 Then
If run(y) = True Then
If d > 29 Then doit = False
Else
If d > 28 Then doit = False
End If

ElseIf m = 4 Or m = 6 Or m = 9 Or m = 11 Then
If d > 30 Then doit = False

ElseIf m > 12 Then
doit = False

Else
If d > 31 Then doit = False

End If

If doit = True Then

If bt2 = 2 Then
bt2 = 3
Timer3.Interval = 20
Timer3.Enabled = True
End If

date1.y = Combo1.Text
date1.m = Combo2.Text
date1.d = Combo3.Text

Open "D:\Myuse\datedistance\date.txt" For Output As #1
Print #1, date1.y & date1.m & date1.d
Close #1

cv = Check1.Value
Open "D:\Myuse\datedistance\data.txt" For Output As #1
Print #1, Mid(Str(cv), 2)
Close #1

Open "D:\Myuse\datedistance\left.txt" For Output As #1
Print #1, Mid(Str(Me.Left), 2)
Close #1

Open "D:\Myuse\datedistance\top.txt" For Output As #1
Print #1, Mid(Str(Me.Top), 2)
Close #1

s = Text1.Text
Label2.Caption = s
Open "D:\Myuse\datedistance\string.txt" For Output As #1
Print #1, s
Close #1

da1 = date1.y & "-" & date1.m & "-" & date1.d
da2 = year(Date) & "-" & Month(Date) & "-" & Day(Date)
Label4.Caption = DateDiff("d", da2, da1)

Else: MsgBox "请输入正确日期"

End If

End Sub

Private Sub Form_DblClick()
If Check1.Visible = False Then
Check1.Visible = True
Else
Check1.Visible = False
End If


End Sub

Private Sub Picture1_MouseDown(Button As Integer, Shift As Integer, X As Single, y As Single)
If bt2 = 2 Then
Timer3.Interval = 20
Timer3.Enabled = True
bt2 = 3
End If

End Sub

Private Sub Picture1_MouseMove(Button As Integer, Shift As Integer, X As Single, y As Single)
If bt2 = 0 Then
bt2 = 1
t2 = 0
h = Form1.Height
v = 10
Timer2.Interval = 20
Timer2.Enabled = True
End If


End Sub

Private Sub Timer1_Timer()
t = t + 1
Debug.Print "1"
pv = pa
Call GetCursorPos(pa)
Form1.Left = Form1.Left + (0 - pv.X + pa.X) * Screen.TwipsPerPixelX      'twip到像素(VB内包含)当窗口属性scalemode =Pixel失效时
Form1.Top = Form1.Top + (0 - pv.y + pa.y) * Screen.TwipsPerPixelY

End Sub
Private Sub Form_MouseDown(Button As Integer, Shift As Integer, X As Single, y As Single)
If Check1.Value = 1 Then
d = 1:
Timer1.Interval = 5: t = 0
Timer1.Enabled = True
'Debug.Print "1"
Call GetCursorPos(pa)
Debug.Print (Str(pa.X))
End If

End Sub
 
Private Sub Form_MouseUp(Button As Integer, Shift As Integer, X As Single, y As Single)
d = 0: Timer1.Interval = 0: Timer1.Interval = 0
End Sub
Private Sub Form_Load()


Me.Width = Screen.Width / 4
Me.Height = Me.Width / 6
Me.Left = Screen.Width - Me.Width - 30 * Screen.TwipsPerPixelX
Me.Top = 15 * Screen.TwipsPerPixelY

Picture1.Top = Me.Height - Picture1.Height
Picture1.Left = (Me.Width - Picture1.Width) / 2

Label1.Top = 5 * Screen.TwipsPerPixelY
Label2.Top = 5 * Screen.TwipsPerPixelY
Label3.Top = 5 * Screen.TwipsPerPixelY
Label4.Top = 5 * Screen.TwipsPerPixelY
Label5.Top = 5 * Screen.TwipsPerPixelY

d = 0
Timer1.Enabled = False
t = 0
Debug.Print Str(Me.Top)

Check1.Value = 0


App.TaskVisible = False
If Dir("D:\Myuse", vbDirectory) = "" Then
MkDir ("D:\Myuse")
End If

If Dir("D:\Myuse\datedistance", vbDirectory) = "" Then
MkDir ("D:\Myuse\datedistance")
Open "D:\Myuse\datedistance\date.txt" For Output As #1
Print #1, "020990101"
Close #1

Open "D:\Myuse\datedistance\data.txt" For Output As #1
Print #1, "0"
Close #1

Open "D:\Myuse\datedistance\left.txt" For Output As #1
Print #1, Mid(Str(Me.Left), 2)
Close #1

Open "D:\Myuse\datedistance\top.txt" For Output As #1
Print #1, Mid(Str(Me.Top), 2)
Close #1

Open "D:\Myuse\datedistance\string.txt" For Output As #1
Print #1, "911"
Close #1

End If

Open "D:\Myuse\datedistance\date.txt" For Input As #1
Input #1, tmp
Close #1
Combo1.Text = Mid(tmp, 1, 5)
Combo2.Text = Mid(tmp, 6, 2)
Combo3.Text = Mid(tmp, 8, 2)
date1.y = Combo1.Text
date1.m = Combo2.Text
date1.d = Combo3.Text

da1 = date1.y & "-" & date1.m & "-" & date1.d
da2 = year(Date) & "-" & Month(Date) & "-" & Day(Date)
Label4.Caption = DateDiff("d", da2, da1)

Open "D:\Myuse\datedistance\data.txt" For Input As #1
Input #1, tmp
Close #1
cv = Val(tmp)
Check1.Value = cv

Open "D:\Myuse\datedistance\left.txt" For Input As #1
Input #1, tmp
Close #1
Me.Left = Val(tmp)

Open "D:\Myuse\datedistance\top.txt" For Input As #1
Input #1, tmp
Close #1
Me.Top = Val(tmp)

Open "D:\Myuse\datedistance\string.txt" For Input As #1
Input #1, tmp
Close #1
s = tmp
Label2.Caption = s
Me.Text1.Text = s
'Me.Text1.Text = Mid(Me.Text1.Text, 2)

For i = 0 To 9999
Combo1.AddItem "0" & Mid(Str(i), 2)
Next i

For i = 1 To 9
Combo2.AddItem "0" & Mid(Str(i), 2)
Next i
For i = 10 To 12
Combo2.AddItem Mid(Str(i), 2)
Next i

Label1.Left = 10 * Screen.TwipsPerPixelX
Label3.Left = (Me.Width - Label3.Width) / 2
Label5.Left = Me.Width - 20 * Screen.TwipsPerPixelX
Label2.Left = (Label3.Left + Label1.Left + Label1.Width - Label2.Width) / 2
Label4.Left = (Label3.Left + Label5.Left + Label3.Width - Label4.Width) / 2

Combo1.Left = 30 * Screen.TwipsPerPixelX
Label6.Left = 35 * Screen.TwipsPerPixelX + Combo1.Width
Combo2.Left = Label6.Left + Label6.Width + Screen.TwipsPerPixelX * 5
Label7.Left = Combo2.Width + Combo2.Left + Screen.TwipsPerPixelX * 5
Combo3.Left = Label7.Left + Label7.Width + Screen.TwipsPerPixelX * 5
Label8.Left = Combo3.Width + Combo3.Left + Screen.TwipsPerPixelX * 5

Dim WshShell As Object
Dim exetemp As String
Set WshShell = CreateObject("wscript.shell")
exetemp = IIf(Right(App.Path, 1) = "\", App.Path, App.Path & "\") & App.EXEName & ".exe"
WshShell.regwrite "HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\Windows\CurrentVersion\Run\", exetemp '加入到注册表，开机运行

End Sub

Private Sub Timer2_Timer()
t2 = t2 + 1
If h < 3701 Then
v = v + a * Timer2.Interval / 80
Else
If Abs(v) > 17 Then
v = 0 - Abs(v - 17)

ElseIf Abs(v) > 13 Then
v = 0 - Abs(v - 13)

ElseIf Abs(v) > 9 Then
v = 0 - Abs(v - 9)

ElseIf Abs(v) > 3 Then
v = 0 - Abs(v - 3)
Else: v = 0: Timer2.Interval = 0: Timer2.Enabled = False: bt2 = 2: Debug.Print "up"
End If

End If

h = h + v * Timer2.Interval / 20
'If t2 * Timer2.Interval > 8600 Then Timer2.Interval = 0: Timer2.Enabled = False:
Form1.Height = h
Picture1.Top = Me.Height - Picture1.Height
If 3700 < h And h < 3700.05 Then Timer2.Interval = 0: Timer2.Enabled = False: bt2 = 2: Debug.Print "up"
End Sub

Private Sub Timer3_Timer()

h = h - 50
If h < Screen.Width / 24 + 30 Then
h = Screen.Width / 24
Timer3.Interval = 0: Timer3.Enabled = False: bt2 = 0
End If
Form1.Height = h
Picture1.Top = Me.Height - Picture1.Height
End Sub
