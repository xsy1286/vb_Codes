VERSION 5.00
Object = "{EAB22AC0-30C1-11CF-A7EB-0000C05BAE0B}#1.1#0"; "ieframe.dll"
Begin VB.Form Form1 
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   1680
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   9180
   LinkTopic       =   "Form1"
   ScaleHeight     =   1680
   ScaleWidth      =   9180
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  '����ȱʡ
   Begin VB.Timer Timerhotkey 
      Left            =   7680
      Top             =   1200
   End
   Begin VB.Timer Timer2 
      Left            =   1680
      Top             =   2400
   End
   Begin VB.Timer Timer1 
      Left            =   960
      Top             =   2400
   End
   Begin ����1.UserControl1 U1 
      Height          =   615
      Left            =   6120
      TabIndex        =   3
      Top             =   1080
      Width           =   975
      _ExtentX        =   1720
      _ExtentY        =   1931
   End
   Begin VB.TextBox Text1 
      BeginProperty Font 
         Name            =   "����"
         Size            =   15
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   560
      Left            =   3240
      TabIndex        =   2
      Text            =   "test��"
      Top             =   1080
      Width           =   2895
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H8000000D&
      Caption         =   "ȡ��"
      Height          =   560
      Left            =   2640
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   1080
      Visible         =   0   'False
      Width           =   615
   End
   Begin SHDocVwCtl.WebBrowser WebBrowser1 
      Height          =   1095
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   9135
      ExtentX         =   16113
      ExtentY         =   1931
      ViewMode        =   0
      Offline         =   0
      Silent          =   0
      RegisterAsBrowser=   0
      RegisterAsDropTarget=   1
      AutoArrange     =   0   'False
      NoClientEdge    =   0   'False
      AlignLeft       =   0   'False
      NoWebView       =   0   'False
      HideFileNames   =   0   'False
      SingleClick     =   0   'False
      SingleSelection =   0   'False
      NoFolders       =   0   'False
      Transparent     =   0   'False
      ViewID          =   "{0057D0E0-3573-11CF-AE69-08002B2E1262}"
      Location        =   "http:///"
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Declare Function GetWindowLong Lib "user32" Alias "GetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long) As Long
Private Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
Private Declare Function SetLayeredWindowAttributes Lib "user32" (ByVal hwnd As Long, ByVal crKey As Long, ByVal bAlpha As Byte, ByVal dwFlags As Long) As Long
Const WS_EX_LAYERED = &H80000
Const GWL_EXSTYLE = (-20)
Const LWA_COLORKEY = &H1
Const Ah = 10   'const ����ͷ��ĸ��д
Const Vh = 240
Dim s As Integer
Dim t As Integer
Dim h As Integer
Dim flag As Boolean


Private Sub Command1_Click()
Text1.Text = ""
If s = 2 Then
Command1.Visible = False
Timer2.Interval = 50
Timer2.Enabled = True
t = 0
s = 3
End If
End Sub

Private Sub Form_Load()

 '�򵥵ļ���ע���ʵ�ֿ�������
Dim WshShell As Object
Dim exetemp As String
Set WshShell = CreateObject("wscript.shell")
exetemp = IIf(Right(App.Path, 1) = "\", App.Path, App.Path & "\") & App.EXEName & ".exe"
WshShell.regwrite "HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\Windows\CurrentVersion\Run\", exetemp
'----------------------------------

Dim transcolor As Long
 transcolor = RGB(66, 66, 66) '���붼66
 Me.BackColor = transcolor '����ı��Լ�����
   Dim rtn As Long
   rtn = GetWindowLong(hwnd, GWL_EXSTYLE)  ' FormX������ֱ�ӵ� hwnd ���� Me.hwnd   ��X = 1,2,3,4 .....)
   rtn = rtn Or WS_EX_LAYERED
   SetWindowLong hwnd, GWL_EXSTYLE, rtn
   SetLayeredWindowAttributes hwnd, transcolor, 0, LWA_COLORKEY
   If Dir("D:\Myuse\2.png") <> "" Then U1.url = "D:\Myuse\2.png"
U1.backsty = 1
Dim r As Long
r = U1.bc(50, 50, 100)
Text1.Left = (Me.Width - Text1.Width) / 2: Text1.Text = ""
Command1.Left = Text1.Left - Command1.Width
U1.Left = Text1.Left + Text1.Width
Me.Width = Me.WebBrowser1.Width
Me.Top = 0
Me.Left = Screen.Width - Me.Width
Me.Height = WebBrowser1.Height + 560
WebBrowser1.Visible = False
WebBrowser1.Top = 0: WebBrowser1.Left = 0
Command1.Visible = False

h = Me.Height
s = 0

  RegisterHotKey Me.hwnd, HotKey_ID, MOD_CONTROL, vbKeyQ  'ע�� Ctrl+ C Ϊ�ȼ�
      HotKey_Flg = False
      Timerhotkey.Interval = 20: Timerhotkey.Enabled = True '�ȼ���timer
End Sub





Private Sub Form_Unload(Cancel As Integer)
 HotKey_Flg = True
       Call UnregisterHotKey(Me.hwnd, HotKey_ID)
End Sub



Private Sub Text1_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then

WebBrowser1.Navigate "http://www.baidu.com/s?wd=" & toAscii(Text1.Text)
WebBrowser1.Visible = True
If s = 0 Then
Timer1.Interval = 50
Timer1.Enabled = True
t = 0
s = 1
End If

End If
End Sub

Private Sub Timer1_Timer()
t = t + 1
If h < (Screen.Height * 2 / 3) Then
h = h + Ah * t
Me.Height = h
WebBrowser1.Height = h - 560
Text1.Top = h - 560: U1.Top = h - 560
Else
Command1.Top = h - 560
Command1.Visible = True
Timer1.Enabled = False
s = 2: t = 0
End If
'Debug.Print CStr(Me.Height)

End Sub

Private Sub Timer2_Timer()
t = t + 1
If h > (1680) Then
h = h - Vh
Me.Height = h
WebBrowser1.Height = h - 560
Text1.Top = h - 560: U1.Top = h - 560
Else
Timer2.Enabled = False
h = 1680
Me.Height = h
WebBrowser1.Height = h - 575: WebBrowser1.Visible = False
Text1.Top = h - 560: U1.Top = h - 560
s = 0: t = 0
End If
'Debug.Print CStr(Me.Height)

End Sub

Private Sub Timerhotkey_Timer()
WaitMessage '�ȴ���Ϣ
          If PeekMessage(Message, Me.hwnd, WM_HOTKEY, WM_HOTKEY, PM_REMOVE) Then '����Ƿ��ȼ�������
         'Me.Hide
          Me.Show
     Text1.SetFocus:
            End If
         DoEvents 'ת�ÿ���Ȩ,�������ϵͳ���������¼�
End Sub

Private Sub U1_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
DoDrag Me
If (GetSystemMetrics(SM_CYSCREEN) * Screen.TwipsPerPixelY - Me.Top) < 9065 Then Me.Top = GetSystemMetrics(SM_CYSCREEN) * Screen.TwipsPerPixelY - 9065
End Sub
Private Sub U1_dbClick()
WebBrowser1.Navigate "http://www.baidu.com/s?wd=" & toAscii(Text1.Text)
WebBrowser1.Visible = True
If s = 0 Then
Timer1.Interval = 50
Timer1.Enabled = True
t = 0
s = 1
End If
End Sub

Function toAscii(sIn As String) As String
 On Error Resume Next
 Dim i As Long
 Dim btmp() As Byte
 btmp = StrConv(sIn, vbFromUnicode)
 For i = LBound(btmp) To UBound(btmp)
 toAscii = toAscii & "%" & Right("00" & Hex(btmp(i)), 2)
 Next
End Function


Private Sub Form_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
DoDrag Me
If (GetSystemMetrics(SM_CYSCREEN) * Screen.TwipsPerPixelY - Me.Top) < 9065 Then Me.Top = GetSystemMetrics(SM_CYSCREEN) * Screen.TwipsPerPixelY - 9065
End Sub



Private Sub WebBrowser1_NewWindow2(ppDisp As Object, Cancel As Boolean) '�����´����д�
On Error Resume Next
Cancel = True
WebBrowser1.Navigate2 (WebBrowser1.Document.activeElement.href)
End Sub

