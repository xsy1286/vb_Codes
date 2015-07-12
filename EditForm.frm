VERSION 5.00
Begin VB.Form EditForm 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "EditForm"
   ClientHeight    =   1065
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   8895
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1065
   ScaleWidth      =   8895
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  '窗口缺省
   Begin VB.Timer Timer1 
      Left            =   2280
      Top             =   600
   End
   Begin VB.VScrollBar VScroll1 
      Height          =   855
      Left            =   6960
      TabIndex        =   11
      Top             =   120
      Width           =   255
   End
   Begin VB.CommandButton Command1 
      Caption         =   "关闭"
      Height          =   855
      Left            =   8520
      TabIndex        =   9
      Top             =   120
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.CheckBox Check1 
      BackColor       =   &H00FFFF80&
      Caption         =   "置顶"
      Height          =   255
      Left            =   2880
      TabIndex        =   5
      Top             =   120
      Width           =   735
   End
   Begin VB.CheckBox Check2 
      BackColor       =   &H00FFFF80&
      Caption         =   "可以移动"
      Height          =   255
      Left            =   5160
      TabIndex        =   4
      Top             =   120
      Width           =   1095
   End
   Begin VB.HScrollBar HScroll1 
      Height          =   255
      Left            =   240
      Max             =   1280
      Min             =   280
      TabIndex        =   3
      Top             =   240
      Value           =   300
      Width           =   2175
   End
   Begin VB.HScrollBar HScroll2 
      Height          =   255
      Left            =   240
      Max             =   800
      Min             =   20
      TabIndex        =   2
      Top             =   0
      Value           =   100
      Width           =   2175
   End
   Begin VB.CheckBox Check3 
      BackColor       =   &H00FFFF80&
      Caption         =   "移动隐藏"
      Height          =   255
      Left            =   3840
      TabIndex        =   1
      Top             =   120
      Width           =   1095
   End
   Begin VB.HScrollBar HScroll3 
      Height          =   255
      Left            =   3960
      Max             =   50
      TabIndex        =   0
      Top             =   600
      Width           =   2295
   End
   Begin VB.Label Label4 
      BackColor       =   &H00FFFF80&
      Caption         =   "亮  度"
      Height          =   615
      Left            =   6480
      TabIndex        =   10
      Top             =   240
      Width           =   255
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FFFF80&
      Caption         =   "    窗口大小"
      Height          =   180
      Left            =   600
      TabIndex        =   8
      Top             =   600
      Width           =   1440
   End
   Begin VB.Label Label2 
      BackColor       =   &H00FFFF80&
      Caption         =   "高 宽"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   11.25
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   495
      Left            =   0
      TabIndex        =   7
      Top             =   30
      Width           =   255
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFF80&
      Caption         =   "移动隐藏时间："
      Height          =   180
      Left            =   2760
      TabIndex        =   6
      Top             =   600
      Width           =   1260
   End
End
Attribute VB_Name = "EditForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'Private Const Wi = 337
'Private Const He = 74
Private Const Appname = "BanSubtitle"
Dim sz(0 To 15) As String

Private Sub Check2_Click()
movable = Check2.Value
'MsgBox CStr(movable)
End Sub

Private Sub Command1_Click()
sz(1) = CStr(Check1.Value)
sz(2) = CStr(Check2.Value)
sz(3) = CStr(Check3.Value)

sz(4) = CStr(Me.HScroll1.Value)
sz(5) = CStr(Me.HScroll2.Value)

sz(6) = CStr(Form1.Left)
sz(7) = CStr(Form1.top)

sz(8) = CStr(HScroll3.Value)
sz(9) = CStr(VScroll1.Value)
Call wr_txtEx(Appname, "data", sz, 10) ':wr_txtEx(

frm1alltop = 1


End Sub

Private Sub Form_Load()


Dim Wi As Integer
Dim He As Integer
'Check1.Value = 0 '刚开始非置顶
Form1.ScaleMode = 3
Wi = GetSystemMetrics(SM_CXSCREEN) / 64
He = GetSystemMetrics(SM_CYSCREEN) / 36

HScroll1.Min = Wi
HScroll2.Min = He

HScroll3.Min = 10 '单位100毫秒
HScroll1.Max = GetSystemMetrics(SM_CXSCREEN)
HScroll2.Max = GetSystemMetrics(SM_CYSCREEN)
HScroll3.Max = 200 '单位100毫秒

VScroll1.Min = 255
VScroll1.Max = 0

init_dir (Appname)
sz(0) = "reserved"
sz(1) = "0": sz(2) = "0": sz(3) = "0":
sz(4) = CStr(Wi): sz(5) = CStr(He)
sz(6) = CStr(Wi * 3): sz(7) = CStr(He * 3)
sz(8) = "50": sz(9) = 105
Call init_txtEx(Appname, "data", sz, 10)
Check1.Value = Val(sz(1))
Check2.Value = Val(sz(2))
Check3.Value = Val(sz(3))

HScroll1.Value = Val(sz(4))
HScroll2.Value = Val(sz(5))


HScroll3.Value = Val(sz(8))
VScroll1.Value = Val(sz(9))

Call rdinitfin
mid_Form Me


If loadOnce <> 8989 Then

loadOnce = 8989

Form1.Left = Val(sz(6))
Form1.top = Val(sz(7))




'top_Form Me, True
Me.Hide

If Check1.Value = 1 Then frm1alltop = 1

Form1.BackColor = RGB(0, VScroll1.Value, 0)
Form1.Show
Form1.Timer2.Interval = 50
End If

'Timer1.Interval = 20
End Sub
Private Sub rdinitfin()
Call Check1_Click
Call Check2_Click
Call Check3_Click

Call HScroll1_Change
Call HScroll2_Change
Call HScroll3_Change

Call VScroll1_Change
End Sub
Private Sub Form_Unload(Cancel As Integer)
Call Command1_Click
If Check3.Value = 1 Then
 Form1.Timer3.Interval = 100
End If

If Check1.Value = 1 Then
 frm1alltop = 1
 Form1.Timer2.Interval = 50: Form1.Timer2.Enabled = True
End If


End Sub

Private Sub HScroll1_Change()
Form1.Width = HScroll1.Value * Screen.TwipsPerPixelX
End Sub

Private Sub HScroll2_Change()
Form1.Height = HScroll2.Value * Screen.TwipsPerPixelY
End Sub

Private Sub Check1_Click()

If Check1.Value = 1 Then
  'frm1alltop = 0:
  top_Form Me, True
 
Else:

top_Form Me, False
End If

End Sub
Private Sub Check3_Click()

If Check3.Value = 1 Then
frs = 1:
Form1.Timer3.Interval = 100 '单位100毫秒

ElseIf Check3.Value = 0 Then

Form1.Timer3.Interval = 0
Form1.Show
End If

End Sub

Private Sub HScroll3_Change()
 p = HScroll3.Value
End Sub

Private Sub Timer1_Timer()
top_Form Me
End Sub

Private Sub VScroll1_Change()
Form1.BackColor = RGB(0, VScroll1.Value, 0)
End Sub
