VERSION 5.00
Begin VB.Form Form2 
   BorderStyle     =   0  'None
   Caption         =   "Form2"
   ClientHeight    =   11475
   ClientLeft      =   2310
   ClientTop       =   1155
   ClientWidth     =   16230
   LinkTopic       =   "Form2"
   Picture         =   "form2.frx":0000
   ScaleHeight     =   11475
   ScaleWidth      =   16230
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton Command5 
      Height          =   180
      Left            =   15480
      TabIndex        =   9
      Top             =   9720
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.CommandButton Command4 
      Height          =   180
      Left            =   14760
      TabIndex        =   8
      Top             =   9720
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.CommandButton Command3 
      Height          =   180
      Left            =   13920
      TabIndex        =   7
      Top             =   9720
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.CheckBox Check1 
      Caption         =   "Check1"
      Height          =   375
      Left            =   4080
      TabIndex        =   6
      Top             =   6600
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.TextBox Text1 
      Height          =   495
      Left            =   2640
      TabIndex        =   5
      Top             =   6600
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Command2"
      Height          =   735
      Left            =   2400
      TabIndex        =   4
      Top             =   4440
      Visible         =   0   'False
      Width           =   1935
   End
   Begin VB.Timer Timer2 
      Left            =   1680
      Top             =   5760
   End
   Begin VB.HScrollBar HScroll1 
      Height          =   255
      Left            =   13800
      Max             =   100
      TabIndex        =   3
      Top             =   9360
      Visible         =   0   'False
      Width           =   2055
   End
   Begin VB.CommandButton Command1 
      Caption         =   "I have known"
      Height          =   735
      Left            =   9480
      TabIndex        =   1
      Top             =   9120
      Width           =   2175
   End
   Begin VB.Timer Timer1 
      Left            =   1680
      Top             =   4920
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   10.5
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   690
      Left            =   11880
      TabIndex        =   2
      Top             =   9120
      Width           =   975
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Please have a rest ,the computer have been turn on for "
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   26.25
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFF80&
      Height          =   1725
      Left            =   1080
      TabIndex        =   0
      Top             =   2040
      Width           =   14850
   End
End
Attribute VB_Name = "Form2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'判断系统是否
'Private Declare Function FindWindow Lib "user32" Alias "FindWindowA" (ByVal lpClassName As String, ByVal lpWindowName As String) As Long

Dim d As Integer

Private Declare Function SendMessage Lib "user32.dll" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lparam As Long) As Long
Private Const WM_APPCOMMAND As Long = &H319
Private Const APPCOMMAND_VOLUME_UP As Long = 10
Private Const APPCOMMAND_VOLUME_DOWN As Long = 9
Private Const APPCOMMAND_VOLUME_MUTE As Long = 8

Private Declare Function SetWindowPos Lib "user32" (ByVal hwnd As Long, ByVal hWndInsertAfter As Long, ByVal X As Long, ByVal Y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long) As Long
Const HWND_TOPMOST = -1
Const HWND_NOTOPMOST = -2
Const SWP_NOSIZE = &H1
Const SWP_NOMOVE = &H2
Const SWP_NOACTIVATE = &H10
Const SWP_SHOWWINDOW = &H40

Const s0 = "Close,"
Const s1 = "I have known"
Sub SetFormTopmost(TheForm As Form)
SetWindowPos Me.hwnd, HWND_TOPMOST, 0, 0, 0, 0, SWP_NOACTIVATE Or SWP_SHOWWINDOW Or SWP_NOMOVE Or SWP_NOSIZE
End Sub

Private Sub Command1_Click()
If Command1.Caption = s0 & s1 Then Unload Form2

End Sub

Private Sub Command2_Click()
End

End Sub

Private Sub Command3_Click()
SendMessage Me.hwnd, WM_APPCOMMAND, &H30292, APPCOMMAND_VOLUME_UP * &H10000

End Sub

Private Sub Command4_Click()
SendMessage Me.hwnd, WM_APPCOMMAND, &H30292, APPCOMMAND_VOLUME_DOWN * &H10000

End Sub

Private Sub Command5_Click()
SendMessage Me.hwnd, WM_APPCOMMAND, &H200EB0, APPCOMMAND_VOLUME_MUTE * &H10000

End Sub

Private Sub Form_Load()
SetWindowPos Me.hwnd, HWND_TOPMOST, 0, 0, 0, 0, SWP_NOMOVE Or SWP_NOSIZE
' 将窗口设为总在最前
Form2.top = 0
Form2.Left = 0

Me.Width = GetSystemMetrics(SM_CXSCREEN) * Screen.TwipsPerPixelX 'screen.width必须用此替代
Me.Height = GetSystemMetrics(SM_CYSCREEN) * Screen.TwipsPerPixelY

SetFormTopmost Me

Open "d:\Myuse\TC\2.txt" For Binary As #1
d = Val(Input(LOF(1), 1))
Close #1

Command1.Caption = s0 & s1 + Str(d)

Command1.Left = (Screen.Width / 2) + 1050

Label2.Left = Command1.Left + Command1.Width
Label2.top = Command1.top + Command1.Height / 2 - 150

HScroll1.Left = Label2.Left + Label2.Width
Command3.Left = HScroll1.Left + 90
Command4.Left = HScroll1.Left + (HScroll1.Width - Command4.Width) / 2
Command5.Left = HScroll1.Left + HScroll1.Width - 90 - Command5.Width
Command3.top = HScroll1.top + 360
Command4.top = HScroll1.top + 360
Command5.top = HScroll1.top + 360

Form2.HScroll1.Value = Form1.w1.settings.volume

Label2.Caption = "": HScroll1.Visible = False: Command3.Visible = False: Command4.Visible = False: Command5.Visible = False
End Sub



Private Sub HScroll1_Change()
Form1.w1.settings.volume = Form2.HScroll1.Value
End Sub


Private Sub Label2_DblClick()

If Label2.Caption = "" Then
Label2.Caption = " Volume:": HScroll1.Visible = True: Command3.Visible = True: Command4.Visible = True: Command5.Visible = True
 Else: Label2.Caption = "": HScroll1.Visible = False: Command3.Visible = False: Command4.Visible = False: Command5.Visible = False
 End If
 
End Sub

Private Sub Timer1_Timer()
Command1.Caption = s0 & s1 + Str(d)
If d = 0 Then Command1.Caption = s0 & s1: Timer1.Interval = 0: Timer2.Interval = 0:: Form1.Show
If d = 1 Then Timer2.Interval = 0: Timer2.Enabled = False
If d > 0 Then d = d - 1


End Sub


Private Sub Form_Paint()
SetFormTopmost Me
End Sub


Private Sub Timer2_Timer()
Form2.Show  '不停的show

End Sub
