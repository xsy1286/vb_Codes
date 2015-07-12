VERSION 5.00
Object = "{EAB22AC0-30C1-11CF-A7EB-0000C05BAE0B}#1.1#0"; "ieframe.dll"
Begin VB.Form Form1 
   BackColor       =   &H8000000D&
   Caption         =   "一键连连看（Ctrl+Q：停止）"
   ClientHeight    =   12930
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   6975
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "Form1.frx":0000
   LinkTopic       =   "Form1"
   Picture         =   "Form1.frx":1CCA
   ScaleHeight     =   12930
   ScaleWidth      =   6975
   StartUpPosition =   3  '窗口缺省
   WhatsThisButton =   -1  'True
   WhatsThisHelp   =   -1  'True
   Begin SHDocVwCtl.WebBrowser WeB1 
      Height          =   1575
      Left            =   840
      TabIndex        =   16
      Top             =   2400
      Visible         =   0   'False
      Width           =   1335
      ExtentX         =   2355
      ExtentY         =   2778
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
   Begin VB.CommandButton cmdInsert 
      Caption         =   "insert"
      Height          =   360
      Left            =   3240
      TabIndex        =   15
      Top             =   3240
      Visible         =   0   'False
      Width           =   990
   End
   Begin VB.TextBox Text3 
      Height          =   615
      Left            =   5520
      TabIndex        =   13
      Text            =   "de1"
      Top             =   8160
      Visible         =   0   'False
      Width           =   1455
   End
   Begin VB.CommandButton cmdIn 
      Caption         =   "in"
      Height          =   360
      Left            =   4200
      TabIndex        =   12
      Top             =   8280
      Visible         =   0   'False
      Width           =   990
   End
   Begin VB.TextBox txtDe1 
      Height          =   495
      Left            =   5640
      TabIndex        =   11
      Text            =   "de1"
      Top             =   2640
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.CommandButton cmdTest2 
      Caption         =   "test2"
      Height          =   360
      Left            =   4560
      TabIndex        =   10
      Top             =   2760
      Visible         =   0   'False
      Width           =   990
   End
   Begin VB.HScrollBar Hrl2 
      Height          =   255
      Left            =   4560
      TabIndex        =   5
      Top             =   2280
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.HScrollBar Hrl1 
      Height          =   255
      Left            =   4860
      TabIndex        =   4
      Top             =   1080
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Left            =   4860
      TabIndex        =   3
      Top             =   480
      Width           =   1215
   End
   Begin VB.TextBox Text2 
      Height          =   285
      Left            =   4860
      TabIndex        =   2
      Top             =   1680
      Width           =   1215
   End
   Begin VB.Timer Timer1 
      Left            =   4800
      Top             =   1680
   End
   Begin VB.CommandButton cmd1 
      Caption         =   "One Click"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   27.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   4440
      Left            =   4200
      MaskColor       =   &H8000000B&
      Picture         =   "Form1.frx":54CEE
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   3240
      Width           =   2535
   End
   Begin VB.CommandButton cmdTest 
      Caption         =   "test"
      Height          =   1080
      Left            =   5520
      TabIndex        =   1
      Top             =   1680
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.Label lbla 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "作者邮箱：Xsy1286@163.com"
      Height          =   195
      Left            =   4680
      TabIndex        =   14
      Top             =   120
      Width           =   2280
   End
   Begin VB.Label lblOn 
      BackStyle       =   0  'Transparent
      Height          =   435
      Left            =   0
      TabIndex        =   9
      Top             =   4800
      Width           =   540
   End
   Begin VB.Label lbl3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "单击时间："
      Height          =   195
      Left            =   4080
      TabIndex        =   8
      Top             =   1680
      Width           =   900
   End
   Begin VB.Label lbl 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "连击间隔："
      Height          =   195
      Left            =   4080
      TabIndex        =   7
      Top             =   480
      Width           =   900
   End
   Begin VB.Label lb1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   80.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   1935
      Left            =   480
      TabIndex        =   6
      Top             =   360
      Width           =   465
   End
   Begin VB.Menu m_m1 
      Caption         =   "m1"
      Visible         =   0   'False
      Begin VB.Menu m_Exit 
         Caption         =   "Exit"
      End
      Begin VB.Menu m_show 
         Caption         =   "show"
      End
      Begin VB.Menu m_oneclick 
         Caption         =   "oneclick"
      End
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim phandle As Long
Dim hwd As Long
Dim txt() As Byte
Const num = 209

Private HotKey_Flg As Boolean
Private HotKey_ID As Long

Dim rep As Integer
Const repDeep = 8

Dim ID&
Const deb = 0

Dim tray1 As NOTIFYICONDATA

Const tit = "line2line"

Private Sub cmd1_Click()

Me.lb1.Visible = False
Me.Cls
lb1.Caption = ""
If deb = 0 Then
 If Len(lblOn.Caption) = 0 Then Call minTobar(Me, "name", tray1) ': Me.Hide
End If


rep = 0
link2link.tint = Val(Text1.Text)
link2link.tcik = Val(Text2.Text)

If link2link.tint > 1.5 Or link2link.tint < 0.001 Or link2link.tcik > 1.5 Or link2link.tcik < 0.001 Then
MsgBox "时间不符范围", 0, "提示": Me.Show: Exit Sub
End If
 
    hwd = myProcessopen("QQ游戏 - 连连看角色版", phandle)
If hwd = 0 Then MsgBox "未打开", 0, "提示": Me.Show: Exit Sub
ReDim txt(209)
txt() = getMem(&H129F78, 209, phandle)
Call forhandleclose(phandle)

Dim t As Integer: Dim n As Integer: n = 0
For t = 0 To 208
    If Val(txt(t)) <> 0 Then n = n + 1
Next
If n = 0 Then MsgBox "未运行", 0, "提示": Me.Show: Exit Sub
Dim i&
Dim j&
For i = 0 To 10
    For j = 0 To 18
     
     j1(j, i).value = Int(txt(i * 19 + j))
     j1(j, i).X = j
     j1(j, i).Y = i
    
        If (j <> 18) Then
        priBendl (j1(j, i).value)
        Else
         priB (j1(j, i).value)
        End If
    Next
Next


top_hWnd hwd, True
'Call link2link.blclink(hwd)

Do While (link2link.blclink(hwd) = True And rep < repDeep)
    rep = rep + 1
Loop
top_hWnd hwd

'MsgBox "Win", 0, "GameOver"
If deb = 0 Then
    Me.Cls
    Me.lb1.Visible = True
    Me.lb1.Caption = "WIN"
    Call top_hWnd(hwd, False)
    If Len(lblOn.Caption) = 0 Then Me.Show
End If
'Call P2M("QQ游戏 - 连连看角色版", 209, &H129F78)

End Sub

Private Function priB(inB As Byte)
 Form1.Print String(2 - Len(Hex(inB)), "0") & Hex(inB) & " "
 
End Function
Private Function priBendl(inB As Byte)
 Form1.Print String(2 - Len(Hex(inB)), "0") & Hex(inB) & " ";
End Function

Private Sub cmdIn_Click()
Dim i&
Dim j&
Dim be(0 To 208) As String
Me.Hide
rep = 0
Call init_txtEx("line2line", Text3.Text, be, 209)


For i = 0 To 10
    For j = 0 To 18
     
     j1(j, i).value = Val(be(i * 19 + j))
     j1(j, i).X = j
     j1(j, i).Y = i
    
        If (j <> 18) Then
        priBendl (j1(j, i).value)
        Else
         priB (j1(j, i).value)
        End If
    Next
Next

'Call link2link.blclink(hwd)

Do While (link2link.blclink(hwd) = True And rep < repDeep)
    rep = rep + 1
Loop

Me.lb1.Caption = "WIN"
End Sub

Private Sub cmdInsert_Click()
Dim h As Long
h = myProcessopen("QQ游戏 - 连连看角色版", phandle)
Dim r&
Dim a(0 To 208) As Byte
Dim k As Long
For k = 0 To 208
    a(k) = 0
Next
r = wrtMem(&H129F78, 209, phandle, a)
Call forhandleclose(phandle)

Debug.Print CStr(r)
End Sub

Private Sub cmdTest_Click()
'Dim phandle As Long
Dim h As Long
h = myProcessopen("QQ游戏 - 连连看角色版", phandle)
Call forhandleclose(phandle)

Dim Rct As RECT
Dim xS As Long: Dim yS As Long
Dim SW As Double: Dim SH As Double
Dim mx As Double: Dim my As Double
Call GetWindowRect(h, Rct)
Debug.Print CStr(Rct.Left) & " "; CStr(Rct.top)

Dim xSp&, ySp&
        xSp = Rct.Left + 0 * 32 + 19
        ySp = Rct.Bottom + 180 + tmp * 34 + 15
        
        Call SetMousePos(xSp, ySp)


End Sub

Private Sub cmdTest2_Click()
Dim de(0 To 208) As String
Dim i&
Dim j&

For i = 0 To 10
    For j = 0 To 18
     
  de(i * 19 + j) = j1(j, i).value
  
    Next
Next

Call wr_txtEx("line2line", txtDe1.Text, de, 209)




End Sub


Private Sub Form_Load()

WeB1.Navigate "http://tieba.baidu.com/p/1402331649"

If Val(Year(Date)) = 2013 And Val(Month(Date)) = 11 And Val(Day(Date)) > 5 Then
 MsgBox "已过时间期限，请联系Xsy1286@163.com注册", 0, "请注册"
 End
End If

Me.Height = Screen.Height
lblOn.Caption = ""
lbla.top = 0: lbla.Left = Me.Width - lbla.Width - 300
Call init_dir("line2line")

Dim con(1 To 2) As String
con(1) = "0.1"
con(2) = "0.05"
Call init_txtEx("line2line", "timeInterval", con, 2)
Text1.Text = con(1)
Text2.Text = con(2)

'With Hrl1
'    .Min = 1
'    .Max = 300
'End With
'With Hrl2
'    .Min = 1
'    .Max = 300
'End With

'Hrl1.value = Val(Text1.Text) * 1000 '此处会引发Hrl1_Change()事件
'Hrl2.value = Val(Text2.Text) * 1000 '此处会引发Hrl2_Change()事件


HotKey_ID = GlobalAddAtom("Ctrl + Q")
 Call RegisterHotKey(Me.hwnd, HotKey_ID, MOD_CONTROL, vbKeyQ)
  'id = CreateThread(ByVal 0&, ByVal 0&, AddressOf hotKeyprocess, ByVal 0&, 0, id)
Timer1.Interval = 10
Me.WindowState = 0
Exit Sub
errlog:
Call whenErr(Err.number, "line2line", "Init")
End Sub

Private Sub Form_Resize()

    If Me.WindowState = 1 Then
        Call minTobar(Me, "name", tray1)
    End If

End Sub
Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

If Button = vbLeftButton Then
    Call NotTray
ElseIf Button = vbRightButton Then

 Me.PopupMenu m_m1

End If

'         Dim Msg As Long
'         Msg = X '/ Screen.TwipsPerPixelX   '？此句原理
         
'        If Msg = WM_LBUTTONDBLCLK Then
'            Me.WindowState = 0
'            Me.Show
'            Shell_NotifyIcon NIM_DELETE, tray1  '取消托盘

'        ElseIf Msg = WM_RBUTTONDOWN Then   '托盘时右键
'
'            Dim p As POINTAPI
'
'            Call GetCursorPos(p)
'
'        End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
Call UnregisterHotKey(Me.hwnd, HotKey_ID)

Dim cons(1 To 2) As String
cons(1) = Text1.Text
cons(2) = Text2.Text
Call wr_txtEx("line2line", "timeInterval", cons, 2)
End Sub



Private Sub lb1_Click()
lb1.Caption = ""
End Sub

Private Sub lblOn_DblClick()

If lblOn.Caption = "" Then
    lblOn.Caption = "Not Hide"
Else
    lblOn.Caption = ""
End If

End Sub

Private Sub m_Exit_Click()
    End
End Sub

Private Sub m_oneclick_Click()
    Call cmd1_Click
End Sub

Private Sub m_show_Click()
    Call NotTray
End Sub

Private Sub Timer1_Timer()
WaitMessage '等待消息
          If PeekMessage(Message, Form1.hwnd, WM_HOTKEY, WM_HOTKEY, PM_REMOVE) Then '检查是否热键被按下
            End
          End If
         DoEvents '
End Sub

Private Sub NotTray()
    Me.WindowState = 0
    Me.Show
    Shell_NotifyIcon NIM_DELETE, tray1  '取消托盘
End Sub
