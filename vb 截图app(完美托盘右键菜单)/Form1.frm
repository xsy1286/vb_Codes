VERSION 5.00
Begin VB.Form Form1 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "截图app        "
   ClientHeight    =   2520
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   5340
   Icon            =   "Form1.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   Picture         =   "Form1.frx":1CCA
   ScaleHeight     =   168
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   356
   StartUpPosition =   3  '窗口缺省
   Begin VB.CommandButton test 
      Caption         =   "test"
      Height          =   615
      Left            =   0
      TabIndex        =   14
      Top             =   1440
      Width           =   375
   End
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
      BackColor       =   &H00FFFFC0&
      Caption         =   "保存剪切板"
      Enabled         =   0   'False
      Height          =   255
      Left            =   3360
      TabIndex        =   11
      Top             =   1320
      Width           =   1575
   End
   Begin VB.CheckBox Check3 
      BackColor       =   &H00FFFFC0&
      Caption         =   "保存默认路径"
      Height          =   255
      Left            =   1680
      TabIndex        =   10
      Top             =   1320
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
      BackColor       =   &H00FFFFC0&
      Caption         =   "截图完成后隐藏"
      Height          =   255
      Left            =   2760
      TabIndex        =   6
      Top             =   840
      UseMaskColor    =   -1  'True
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
      BackColor       =   &H00FFFFC0&
      Caption         =   "使用快捷键"
      Height          =   255
      Left            =   840
      TabIndex        =   2
      Top             =   840
      Value           =   1  'Checked
      Width           =   1575
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H00FFFFC0&
      Caption         =   "点击截图"
      Height          =   450
      Left            =   1800
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   120
      Width           =   1525
   End
   Begin VB.Timer Timer1 
      Left            =   600
      Top             =   2520
   End
   Begin VB.PictureBox Picture1 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      Height          =   1215
      Left            =   1080
      ScaleHeight     =   77
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   125
      TabIndex        =   0
      Top             =   2640
      Visible         =   0   'False
      Width           =   1935
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "双击后:"
      Height          =   375
      Left            =   840
      TabIndex        =   9
      Top             =   1275
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "欢迎任何建议及反馈函至作者邮箱:  Xsy1286@163.com"
      Height          =   375
      Left            =   600
      TabIndex        =   8
      Top             =   2160
      Width           =   4455
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "My Company -v3.3"
      Height          =   180
      Left            =   3600
      TabIndex        =   3
      Top             =   240
      Width           =   1440
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit


Private Tray As NOTIFYICONDATA
Dim temp As String
Dim tmp As String
Dim r As Integer
#Const stp = 0
Const appName = "shot"
Dim hId As Long, HNum As Integer

Private Sub Check1_Click()
Dim r As Long
If Check1.Value = 1 Then
Form1.Text1.Text = "1"
  r = insertHotKey(Me.hwnd, hId, True, False, False, vbKeyS, HNum) '注册 Ctrl+ S 为热键
      HotKey_Flg = False
      
#If stp <> 0 Then
  Stop
#End If

Else:
Form1.Text1.Text = "0"
 HotKey_Flg = True
     r = unHotKey(Me.hwnd, hId)
     
#If stp <> 0 Then
  Stop
#End If

End If
Open "d:\Myuse\shot\address.txt" For Output As #1
Print #1, (Form1.Text1.Text) & (Form1.Text2.Text) & (Form1.Text3.Text) & CStr(dfaddress) & Mid(Str(clip), 2) & Mid(Str(v6), 2)
Close #1
Debug.Print r
End Sub

Private Sub Check2_Click() 'check2是和Form1.Text3.Text关联
'
Form1.Text3.Text = CStr(Check2.Value)
Open "d:\Myuse\shot\address.txt" For Output As #1
Print #1, (Form1.Text1.Text) & (Form1.Text2.Text) & (Form1.Text3.Text) & CStr(dfaddress) & Mid(Str(clip), 2) & Mid(Str(v6), 2)
Close #1


End Sub
Private Sub Check3_Click() 'check3是和dfaddress关联

Select Case Form1.Combo1.Text
Case ".jpg"
v6 = 1
Case ".png"
v6 = 2
Case ".bmp"
v6 = 3
End Select

'Open "d:\Myuse\shot\address.txt" For Output As #1
'Print #1, (Form1.Text1.Text) & (Form1.Text2.Text) & (Form1.Text3.Text) & CStr(dfaddress) & Mid(Str(clip), 2) & Mid(Str(v6), 2)
'Close #1


If Check3.Value = 1 Then



    If Dir(Form1.Text4.Text) <> "" Then
    
        Open "d:\Myuse\shot\address2.txt" For Output As #1
        Print #1, tmp
        Close #1
        dfaddress = 1
        
        Form1.Text4.Enabled = False: Form1.Combo1.Enabled = False
        Else:
        Check3.Value = 0
        MsgBox "路径错误，请重新输入": Exit Sub
        
    End If

Else:
    dfaddress = 0
    Form1.Text4.Enabled = True: Form1.Combo1.Enabled = True
End If

Open "d:\Myuse\shot\address.txt" For Output As #1
Print #1, (Form1.Text1.Text) & (Form1.Text2.Text) & (Form1.Text3.Text) & Str(dfaddress) & Mid(Str(clip), 2) & Mid(Str(v6), 2)
Close #1
End Sub

Private Sub Check4_Click()
If Check4.Value = 1 Then '变化显示后，才此Sub
clip = 1
Else:
clip = 0
End If

Open "d:\Myuse\shot\address.txt" For Output As #1
Print #1, (Form1.Text1.Text) & (Form1.Text2.Text) & (Form1.Text3.Text) & CStr(dfaddress) & Mid(Str(clip), 2) & Mid(Str(v6), 2)
Close #1
End Sub



'当双击托盘时恢复原状
Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Dim Msg As Long
Msg = X '/ Screen.TwipsPerPixelX
If Msg = WM_LBUTTONDBLCLK Then
Me.WindowState = 0
Me.Show: iftray = 0
'Shell_NotifyIcon NIM_DELETE, Tray   '取消托盘
Unload Form3
End If
If Msg = WM_RBUTTONDOWN Then   '托盘时右键
Dim p As POINTAPI
Call GetCursorPos(p)
'Debug.Print "point"
'Debug.Print Str(p.x)
'Debug.Print Str(p.y)

tx = p.X
ty = p.Y
Load Form4



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

Unload Me
End Sub

Private Sub Form_Load()

If m_debug = 1 Then
  test.Visible = True
Else
 test.Visible = False
End If

Text4.Left = Check1.Left
'Command1.Left = (Screen.Width - Command1.Width) / 2
Tray.cbSize = Len(Tray)
Tray.uID = vbNull
Tray.hwnd = Me.hwnd
Tray.uFlags = NIF_TIP Or NIF_MESSAGE Or NIF_ICON
Tray.uCallbackMessage = WM_MOUSEMOVE
Tray.hIcon = Me.Icon
Tray.szTip = "截图app" & vbNullChar
Shell_NotifyIcon NIM_ADD, Tray
Me.Hide: iftray = 1

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
dfaddress = Val(Mid(temp, 4, 1))
clip = Val(Mid(temp, 5, 1))
Form1.Check1.Value = Val(Mid(temp, 1, 1))
Form1.Check2.Value = Val(Mid(temp, 3, 1))
Form1.Check3.Value = dfaddress
Form1.Check4.Value = clip
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

Call setStartUp(True, appName)

If Text1.Text = "1" Then
   Call insertHotKey(Me.hwnd, hId, True, False, False, vbKeyS, HNum)
      HotKey_Flg = False
  '  Debug.Print r
End If

 Timer1.Interval = 2
  
Call setAttribute(Me.hwnd, Me.BackColor, 208, 2)

End Sub

Private Sub Form_Unload(Cancel As Integer)
Shell_NotifyIcon NIM_DELETE, Tray   '取消托盘
Load Form3
 iftray = 0
 HotKey_Flg = True
       Call unHotKey(Me.hwnd, hId)
End Sub

'
Private Sub Timer1_Timer()
WaitMessage '等待消息
          If PeekMessage(Message, Me.hwnd, WM_HOTKEY, WM_HOTKEY, PM_REMOVE) Then '检查是否热键被按下
        Form1.Hide

        If (0) Then  'Maybe可以去掉 it's for the  reason that Form1显示时无托盘图标
Tray.cbSize = Len(Tray)
Tray.uID = vbNull
Tray.hwnd = Me.hwnd
Tray.uFlags = NIF_TIP Or NIF_MESSAGE Or NIF_ICON
Tray.uCallbackMessage = WM_MOUSEMOVE
Tray.hIcon = Me.Icon
Tray.szTip = "截图app" & vbNullChar
End If

'Shell_NotifyIcon NIM_ADD, Tray  '可能是每次截图后多个托盘图标的原因

Load Form2
          End If
         DoEvents '转让控制权,允许操作系统处理其他事件

End Sub
