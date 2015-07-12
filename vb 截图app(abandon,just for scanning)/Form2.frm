VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "ComDlg32.OCX"
Begin VB.Form Form2 
   AutoRedraw      =   -1  'True
   BorderStyle     =   0  'None
   Caption         =   "Form2"
   ClientHeight    =   5160
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   8880
   LinkTopic       =   "Form2"
   ScaleHeight     =   344
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   592
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  '窗口缺省
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   5280
      Top             =   2880
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.CheckBox Check1 
      Alignment       =   1  'Right Justify
      BackColor       =   &H8000000B&
      Caption         =   "完成后打开图片"
      Height          =   255
      Left            =   960
      TabIndex        =   11
      Top             =   60
      Width           =   1695
   End
   Begin VB.ComboBox Combo1 
      Height          =   300
      Left            =   4680
      TabIndex        =   9
      Text            =   ".bmp"
      Top             =   360
      Visible         =   0   'False
      Width           =   900
   End
   Begin VB.TextBox Text2 
      Height          =   375
      Left            =   960
      TabIndex        =   8
      Top             =   960
      Visible         =   0   'False
      Width           =   6000
   End
   Begin VB.CommandButton Command5 
      BackColor       =   &H00FFFF80&
      Caption         =   "保存"
      Height          =   300
      Left            =   5760
      Style           =   1  'Graphical
      TabIndex        =   7
      Top             =   2160
      Visible         =   0   'False
      Width           =   500
   End
   Begin VB.TextBox Text1 
      Height          =   375
      Left            =   2640
      TabIndex        =   5
      Top             =   360
      Visible         =   0   'False
      Width           =   1695
   End
   Begin VB.DirListBox Dir1 
      Height          =   1770
      Left            =   960
      TabIndex        =   4
      Top             =   2280
      Visible         =   0   'False
      Width           =   4095
   End
   Begin VB.CommandButton Command4 
      BackColor       =   &H00FFFF00&
      Caption         =   "取消"
      Height          =   300
      Left            =   6960
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   2880
      Visible         =   0   'False
      Width           =   500
   End
   Begin VB.CommandButton Command3 
      BackColor       =   &H00FFFF00&
      Caption         =   "确定"
      Height          =   300
      Left            =   6480
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   2880
      Visible         =   0   'False
      Width           =   500
   End
   Begin VB.CommandButton Command2 
      BackColor       =   &H00FFFFFF&
      Caption         =   "取消截图"
      Height          =   375
      Left            =   0
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   0
      Width           =   855
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "关闭"
      Height          =   375
      Left            =   0
      MaskColor       =   &H000080FF&
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   360
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.Label Label2 
      Caption         =   "Label2"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   14.25
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1215
      Left            =   480
      TabIndex        =   10
      Top             =   1680
      Visible         =   0   'False
      Width           =   2535
   End
   Begin VB.Label Label1 
      Caption         =   "地址及名称"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   14.25
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   960
      TabIndex        =   6
      Top             =   360
      Visible         =   0   'False
      Width           =   1575
   End
   Begin VB.Shape Shape1 
      BorderColor     =   &H00000000&
      BorderWidth     =   2
      Height          =   15
      Left            =   1200
      Top             =   1080
      Width           =   15
   End
End
Attribute VB_Name = "Form2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim x1, x2, y1, y2 As Long

Private Declare Function SetWindowPos Lib "user32" (ByVal hWnd As Long, ByVal hWndInsertAfter As Long, ByVal X As Long, ByVal Y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long) As Long
Const HWND_TOPMOST = -1
Const HWND_NOTOPMOST = -2
Const SWP_NOSIZE = &H1
Const SWP_NOMOVE = &H2
Const SWP_NOACTIVATE = &H10
Const SWP_SHOWWINDOW = &H40

Sub SetFormTopmost(TheForm As Form)
SetWindowPos Me.hWnd, HWND_TOPMOST, 0, 0, 0, 0, SWP_NOACTIVATE Or SWP_SHOWWINDOW Or SWP_NOMOVE Or SWP_NOSIZE
End Sub


Private Sub Check1_Click()
If Check1.Value = 1 Then
Form1.Text2.Text = "1"
Else:
Form1.Text2.Text = "0"
End If
Open "d:\Myuse\shot\address.txt" For Output As #1
Print #1, (Form1.Text1.Text) & (Form1.Text2.Text) & (Form1.Text3.Text) & Mid(Str(v4), 2) & Mid(Str(v5), 2) & Mid(Str(v6), 2)
Close #1
End Sub



Private Sub Combo1_Change()
Text2.Text = Dir1.Path & "\" & Text1.Text & Combo1.Text
End Sub

Private Sub Command1_Click()
Shape1.Visible = False
Call UnHooK

End

End Sub

Private Sub Command2_Click()

Call UnHooK
If v = 0 Then
Form1.Show
Form1.Refresh
Form1.Check1.Refresh
Form1.Check2.Refresh
Form1.Check3.Refresh
Form1.Check4.Refresh
Form1.Command1.Refresh
Form1.Label1.Refresh
Form1.Label2.Refresh
Form1.Label3.Refresh
Form1.Text4.Refresh
Form1.Combo1.Refresh
End If
 Unload Form2
End Sub

Private Sub Command3_Click()
Form1.Picture1.Refresh
Form2.Text1.Text = Year(Date) & Month(Date) & Day(Date) & Hour(time) & Minute(time) & Second(time) & "截图"
 Call UnHooK
Form2.Shape1.Visible = False
Form2.Shape1.Width = 0
Form2.Shape1.Height = 0
Form2.Command3.Visible = False
Form2.Command4.Visible = False
Form1.Picture1.Refresh

CommonDialog1.ShowSave

If Len(Form2.CommonDialog1.FileName) <> 0 Then
SavePicture Form1.Picture1.Image, (Form2.CommonDialog1.FileName)

If v = 0 Then
Form1.Show
Form1.Refresh
Form1.Check1.Refresh
Form1.Check2.Refresh
Form1.Check3.Refresh
Form1.Check4.Refresh
Form1.Command1.Refresh
Form1.Label1.Refresh
Form1.Label2.Refresh
Form1.Label3.Refresh
Form1.Text4.Refresh
Form1.Combo1.Refresh
End If

Form2.Hide

 
                 If Form2.Check1.Value = 1 Then Shell "explorer " & Form2.CommonDialog1.FileName, 1  'shell函数，打开C盘  Shell "explorer 路径"
 
 Unload Form2
Else:
Call HooK
 
 End If
 
End Sub

Private Sub Command4_Click()
d = 0
Shape1.Width = 0
Shape1.Height = 0
Shape1.Visible = False
Command3.Visible = False
Command4.Visible = False
dn = 0
End Sub

Private Sub Dir1_Change()
Text2.Text = Dir1.Path & "\" & Text1.Text & Combo1.Text
End Sub


Private Sub Form_Load()
SetWindowPos Form2.hWnd, HWND_TOPMOST, 0, 0, 0, 0, SWP_NOMOVE Or SWP_NOSIZE
' 将窗口设为总在最前
SetFormTopmost Form2
CommonDialog1.Filter = "Bmp Files (*.BMP)|*.bmp|Jpg Files (*.JPG)|*.jpg|Png Files (*.PNG)|*.png|All Files (*.*)|*.*"

Form2.Left = 0
Form2.Top = 0
Form2.Width = Screen.Width
Form2.Height = Screen.Height
BitBlt Form2.hdc, 0, 0, Screen.Width, Screen.Height, GetDC(0), 0, 0, vbSrcCopy
Command3.Visible = False
Shape1.Visible = False

Command2.Visible = True
Command4.Visible = False
Text1.Visible = False
Text2.Visible = False
Dir1.Visible = False
Label1.Visible = False
Command5.Visible = False

'用户自定义

Combo1.Visible = False
Combo1.AddItem (".jpg")
Combo1.AddItem (".png")
Combo1.AddItem (".bmp")

'防止在form2 load前 提到form2物件 变量而导致form2 load 及 form2.show
Form2.Check1.Value = Val(Form1.Text2.Text)
dn = 0
tt = 0

Call HooK
End Sub
Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 27 Then Call Cancelshoot
End Sub
Private Sub Check1_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 27 Then Call Cancelshoot
End Sub
Private Sub Command3_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 27 Then Call Cancelshoot
End Sub
Private Sub Command4_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 27 Then Call Cancelshoot
End Sub

Private Sub Command2_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 27 Then Call Cancelshoot
End Sub
Private Sub Form_Unload(Cancel As Integer)

 Call UnHooK
End Sub
Private Sub Text1_Change()
Text2.Text = Dir1.Path & "\" & Text1.Text & Combo1.Text
End Sub


Private Sub Label1_DblClick()
If dn = 1 Then Call dl
End Sub
Private Sub Label2_DblClick()
If dn = 1 Then Call dl
End Sub
Private Sub Form_DblClick()
If dn = 1 Then Call dl
End Sub
Private Sub Text1_DblClick()
If dn = 1 Then Call dl
End Sub

Private Sub Text2_DblClick()
If dn = 1 Then Call dl
End Sub
