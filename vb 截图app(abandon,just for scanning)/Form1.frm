VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "截图app        "
   ClientHeight    =   2535
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   4980
   Icon            =   "Form1.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   169
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   332
   StartUpPosition =   3  '窗口缺省
   Begin VB.ComboBox Combo1 
      Height          =   300
      Left            =   4080
      TabIndex        =   13
      Text            =   ".bmp"
      Top             =   1590
      Width           =   900
   End
   Begin VB.TextBox Text4 
      Enabled         =   0   'False
      Height          =   300
      Left            =   840
      TabIndex        =   12
      Top             =   1590
      Width           =   3015
   End
   Begin VB.CheckBox Check4 
      Caption         =   "保存剪切板"
      Height          =   375
      Left            =   2760
      TabIndex        =   11
      Top             =   1200
      Width           =   1575
   End
   Begin VB.CheckBox Check3 
      Caption         =   "保存默认路径"
      Height          =   255
      Left            =   840
      TabIndex        =   10
      Top             =   1230
      Width           =   1575
   End
   Begin VB.TextBox Text3 
      Height          =   270
      Left            =   360
      TabIndex        =   7
      Top             =   480
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.CheckBox Check2 
      Caption         =   "截图完成后隐藏"
      Height          =   375
      Left            =   2760
      TabIndex        =   6
      Top             =   720
      Width           =   1575
   End
   Begin VB.TextBox Text2 
      Height          =   270
      Left            =   1200
      TabIndex        =   5
      Top             =   480
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.TextBox Text1 
      Height          =   270
      Left            =   1080
      TabIndex        =   4
      Top             =   120
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.CheckBox Check1 
      Caption         =   "使用快捷键"
      Height          =   375
      Left            =   840
      TabIndex        =   2
      Top             =   720
      Value           =   1  'Checked
      Width           =   1455
   End
   Begin VB.CommandButton Command1 
      Caption         =   "点击截图"
      Height          =   450
      Left            =   1610
      TabIndex        =   1
      Top             =   120
      Width           =   1525
   End
   Begin VB.Timer Timer1 
      Left            =   1800
      Top             =   2400
   End
   Begin VB.PictureBox Picture1 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      Height          =   1215
      Left            =   1200
      ScaleHeight     =   77
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   125
      TabIndex        =   0
      Top             =   2400
      Visible         =   0   'False
      Width           =   1935
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "双击后:"
      Height          =   375
      Left            =   120
      TabIndex        =   9
      Top             =   1275
      Width           =   735
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "欢迎任何建议及反馈函至作者邮箱:  Xsy1286@163.com"
      Height          =   375
      Left            =   360
      TabIndex        =   8
      Top             =   2160
      Width           =   4455
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "My Company -0.9"
      Height          =   495
      Left            =   3600
      TabIndex        =   3
      Top             =   240
      Width           =   1935
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Declare Function Shell_NotifyIcon Lib "shell32.dll" (ByVal dwMessage As Long, lpData As NOTIFYICONDATA) As Long

Const NIM_ADD = &H0                     '在任务栏中增加一个图标
Const NIM_DELETE = &H2                  '删除任务栏中的一个图标
Const NIM_MODIFY = &H1                  '修改任务栏中个图标信息
Const NIF_ICON = &H2                    '
Const NIF_MESSAGE = &H1                 'NOTIFYICONDATA结构中uFlags的控制信息
Const NIF_TIP = &H4                     '
Const WM_MOUSEMOVE = &H200              '
Const WM_LBUTTONDBLCLK = &H203          '
 Const MOD_ALT = &H1
Const MOD_CONTROL = &H2
 Const MOD_SHIFT = &H4
Const PM_REMOVE = &H1
 Const WM_HOTKEY = &H312
Private Type NOTIFYICONDATA
  cbSize As Long                        '该数据结构的大小
  hwnd As Long                          '处理任务栏中图标的窗口句柄
  uID As Long                           '定义的任务栏中图标的标识
  uFlags As Long                        '任务栏图标功能控制，可以是以下值的组合（一般全包括）
                                        'NIF_MESSAGE 表示发送控制消息；
                                        'NIF_ICON表示显示控制栏中的图标；
                                        'NIF_TIP表示任务栏中的图标有动态提示。
  uCallbackMessage As Long '任务栏图标通过它与用户程序交换消息，处理该消息的窗口由hWnd决定
  hIcon As Long '任务栏中的图标的控制句柄
  szTip As String * 64 '图标的提示信息
End Type
Dim Tray As NOTIFYICONDATA
Dim temp As String
Dim tmp As String
Dim r As Integer

Private Sub Check1_Click()
If Check1.Value = 1 Then
Form1.Text1.Text = "1"
  RegisterHotKey Me.hwnd, HotKey_ID, MOD_CONTROL, vbKeyS  '注册 Ctrl+Alt+ S 为热键
      HotKey_Flg = False
Else:
Form1.Text1.Text = "0"
 HotKey_Flg = True
       Call UnregisterHotKey(Me.hwnd, HotKey_ID)
End If
Open "d:\Myuse\shot\address.txt" For Output As #1
Print #1, (Form1.Text1.Text) & (Form1.Text2.Text) & (Form1.Text3.Text) & Mid(Str(v4), 2) & Mid(Str(v5), 2) & Mid(Str(v6), 2)
Close #1
End Sub

Private Sub Check2_Click()
If Check2.Value = 1 Then
Form1.Text3.Text = "1"
  RegisterHotKey Me.hwnd, HotKey_ID, MOD_CONTROL, vbKeyS  '注册 Ctrl+Alt+ S 为热键
      HotKey_Flg = False
Else:
Form1.Text3.Text = "0"
 HotKey_Flg = True
       Call UnregisterHotKey(Me.hwnd, HotKey_ID)
End If
v = Val(Text3.Text)
Open "d:\Myuse\shot\address.txt" For Output As #1
Print #1, (Form1.Text1.Text) & (Form1.Text2.Text) & (Form1.Text3.Text) & Mid(Str(v4), 2) & Mid(Str(v5), 2) & Mid(Str(v6), 2)
Close #1
End Sub
Private Sub Check3_Click()

Select Case Form1.Combo1.Text
Case ".jpg"
v6 = 1
Case ".png"
v6 = 2
Case ".bmp"
v6 = 3
End Select

Open "d:\Myuse\shot\address.txt" For Output As #1
Print #1, (Form1.Text1.Text) & (Form1.Text2.Text) & (Form1.Text3.Text) & Mid(Str(v4), 2) & Mid(Str(v5), 2) & Mid(Str(v6), 2)
Close #1

If Check3.Value = 1 Then



If Dir(Form1.Text4.Text) <> "" Then

Open "d:\Myuse\shot\address2.txt" For Output As #1
Print #1, tmp
Close #1
v4 = 1

Form1.Text4.Enabled = False: Form1.Combo1.Enabled = False
Else:
Check3.Value = 0
MsgBox "路径错误，请重新输入"
End If

Else:
v4 = 0
Form1.Text4.Enabled = True: Form1.Combo1.Enabled = True
End If

Open "d:\Myuse\shot\address.txt" For Output As #1
Print #1, (Form1.Text1.Text) & (Form1.Text2.Text) & (Form1.Text3.Text) & Mid(Str(v4), 2) & Mid(Str(v5), 2) & Mid(Str(v6), 2)
Close #1
End Sub

Private Sub Check4_Click()
If Check4.Value = 1 Then
v5 = 1
Else:
v5 = 0
End If

Open "d:\Myuse\shot\address.txt" For Output As #1
Print #1, (Form1.Text1.Text) & (Form1.Text2.Text) & (Form1.Text3.Text) & Mid(Str(v4), 2) & Mid(Str(v5), 2) & Mid(Str(v6), 2)
Close #1
End Sub

Private Sub Form_DblClick()

End Sub

'当双击托盘时恢复原状
Private Sub Form_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
Dim Msg As Long
Msg = x '/ Screen.TwipsPerPixelX
If Msg = WM_LBUTTONDBLCLK Then
Me.WindowState = 0
Me.Show
Shell_NotifyIcon NIM_DELETE, Tray
Unload Form3
End If
If Msg = WM_RBUTTONDOWN Then   '托盘时右键
Dim p As POINTAPI
Call GetCursorPos(p)
'Debug.Print "point"
'Debug.Print Str(p.x)
'Debug.Print Str(p.y)

tx = p.x
ty = p.y
'Load FormRightMemu  //可使用windows托盘右键菜单



End If
End Sub
'当点最小化是为托盘状态
Private Sub Form_Resize()
If Me.WindowState = 1 Then
Tray.cbSize = Len(Tray)
Tray.uID = vbNull
Tray.hwnd = Me.hwnd
Tray.uFlags = NIF_TIP Or NIF_MESSAGE Or NIF_ICON
Tray.uCallbackMessage = WM_MOUSEMOVE
Tray.hIcon = Me.Icon
Tray.szTip = "截图app" & vbNullChar
Shell_NotifyIcon NIM_ADD, Tray
Me.Hide
End If
End Sub


Private Sub Command1_Click()
Form1.Hide
Tray.cbSize = Len(Tray)
Tray.uID = vbNull
Tray.hwnd = Me.hwnd
Tray.uFlags = NIF_TIP Or NIF_MESSAGE Or NIF_ICON
Tray.uCallbackMessage = WM_MOUSEMOVE
Tray.hIcon = Me.Icon
Tray.szTip = "截图app" & vbNullChar
Shell_NotifyIcon NIM_ADD, Tray
Load Form2
Form2.Show

End Sub

Private Sub Form_Load()

Text4.Left = Check1.Left

Tray.cbSize = Len(Tray)
Tray.uID = vbNull
Tray.hwnd = Me.hwnd
Tray.uFlags = NIF_TIP Or NIF_MESSAGE Or NIF_ICON
Tray.uCallbackMessage = WM_MOUSEMOVE
Tray.hIcon = Me.Icon
Tray.szTip = "截图app" & vbNullChar
Shell_NotifyIcon NIM_ADD, Tray
Me.Hide

  Form1.Left = (Screen.Width - Form1.Width) / 2
Form1.Top = (Screen.Height - Form1.Height) / 2
App.TaskVisible = False
If Dir("D:\Myuse", vbDirectory) = "" Then
MkDir ("D:\Myuse")
End If
If Dir("D:\Myuse\shot", vbDirectory) = "" Then
MkDir ("D:\Myuse\shot")
Open "d:\Myuse\shot\address.txt" For Output As #1
Print #1, "100111"
Close #1

Open "d:\Myuse\shot\address2.txt" For Output As #1
Print #1, "D:\Myuse\shot\"
Close #1

Open "d:\Myuse\shot\说明.txt" For Output As #1
Print #1, "本软件可开机自动启动，360可能会拦截报错。如需取消开机自动启动，请在360开机启动项中删除"
Close #1
Open "d:\Myuse\shot\说明.txt" For Append As #1
Print #1, "截图快捷键为 CTRL + S,最小化及关闭按钮都为最小化至托盘,如需关闭软件请打开任务管理器关闭"
Close #1
Open "d:\Myuse\shot\说明.txt" For Append As #1
Print #1, "反馈至作者邮箱 :Xsy1286@163.com"
Close #1
Shell "notepad d:\Myuse\shot\说明.txt", vbNormalFocus
End If

Open "d:\Myuse\shot\address2.txt" For Input As #1
Input #1, tmp
Close #1
'tmp = "D:\Myuse\shot\"
Form1.Text4.Text = tmp

Combo1.AddItem (".jpg")
Combo1.AddItem (".png")
Combo1.AddItem (".bmp")

Open "d:\Myuse\shot\address.txt" For Binary As #1
temp = Input(LOF(1), 1)
Close #1
Text1.Text = Mid(temp, 1, 1)
Text2.Text = Mid(temp, 2, 1)
Text3.Text = Mid(temp, 3, 1)
v4 = Val(Mid(temp, 4, 1))
v5 = Val(Mid(temp, 5, 1))
Form1.Check1.Value = Val(Mid(temp, 1, 1))
Form1.Check2.Value = Val(Mid(temp, 3, 1))
Form1.Check3.Value = v4
Form1.Check4.Value = v5
v = Form1.Check2.Value
Dim tm As String
tm = Mid(temp, 6, 1)
If tm = "1" Then Combo1.Text = ".jpg"
If tm = "2" Then Combo1.Text = ".png"
If tm = "3" Then Combo1.Text = ".bmp"

If Form1.Check3.Value = 1 Then
Form1.Text4.Enabled = False: Form1.Combo1.Enabled = False
Else: Form1.Text4.Enabled = True: Form1.Combo1.Enabled = True
End If


 '简单的加入注册表实现开机启动

Dim WshShell As Object
Dim exetemp As String
Set WshShell = CreateObject("wscript.shell")
exetemp = IIf(Right(App.Path, 1) = "\", App.Path, App.Path & "\") & App.EXEName & ".exe"
WshShell.regwrite "HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\Windows\CurrentVersion\Run\", exetemp '加入到注册表，开机运行
HotKey_ID = GlobalAddAtom("Ctrl + S")
       'RegisterHotKey Me.hWnd, &HBFFF&, MOD_CONTROL + MOD_ALT, vbKeyG '注册 Ctrl+Alt+ G 为热键
      If Text1.Text = "1" Then
     r = RegisterHotKey(Me.hwnd, HotKey_ID, MOD_CONTROL, vbKeyS)   '注册 Ctrl+ S 为热键
      HotKey_Flg = False
  '  Debug.Print r
End If

  Timer1.Interval = 5
End Sub

Private Sub Form_Unload(Cancel As Integer)
Load Form3

 HotKey_Flg = True
       Call UnregisterHotKey(Me.hwnd, HotKey_ID)
End Sub

Private Sub Timer1_Timer()
WaitMessage '等待消息
          If PeekMessage(Message, Me.hwnd, WM_HOTKEY, WM_HOTKEY, PM_REMOVE) Then '检查是否热键被按下
        Form1.Hide
Tray.cbSize = Len(Tray)
Tray.uID = vbNull
Tray.hwnd = Me.hwnd
Tray.uFlags = NIF_TIP Or NIF_MESSAGE Or NIF_ICON
Tray.uCallbackMessage = WM_MOUSEMOVE
Tray.hIcon = Me.Icon
Tray.szTip = "截图app" & vbNullChar
Shell_NotifyIcon NIM_ADD, Tray

Form2.Show
          End If
         DoEvents '转让控制权,允许操作系统处理其他事件
      
End Sub
