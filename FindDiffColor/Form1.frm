VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   7215
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   9585
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   ScaleHeight     =   7215
   ScaleWidth      =   9585
   StartUpPosition =   3  '窗口缺省
   Begin VB.Timer Timer1 
      Left            =   5160
      Top             =   5040
   End
   Begin 工程1.Value Value1 
      Height          =   2535
      Left            =   5280
      TabIndex        =   3
      Top             =   1440
      Width           =   3615
      _ExtentX        =   6376
      _ExtentY        =   4471
   End
   Begin VB.PictureBox U1 
      Height          =   2775
      Left            =   480
      ScaleHeight     =   2715
      ScaleWidth      =   3435
      TabIndex        =   2
      Top             =   3960
      Width           =   3495
   End
   Begin VB.CommandButton cmd2 
      Caption         =   "2"
      Height          =   2040
      Left            =   1080
      TabIndex        =   1
      Top             =   960
      Width           =   1095
   End
   Begin VB.CommandButton cmd1 
      Caption         =   "1"
      Height          =   2160
      Left            =   2760
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   840
      Width           =   1335
   End
   Begin VB.Label Label1 
      Caption         =   "Label1"
      Height          =   375
      Left            =   1320
      TabIndex        =   4
      Top             =   240
      Width           =   855
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim hwd As Long
Dim ph As Long
Dim dc As Long

Dim Rct As RECT

Const tcik = 0.1

Const tit = "Find Difference"


Dim s(0 To 1) As Integer
Dim step As Integer
Const getHWnd = 1

    Dim ci() As Long
    Dim n As Integer
Dim hi As Integer
Dim wi As Integer



Private Sub cmd1_Click()
    hwd = myProcessopen("大家来找茬", ph)

    If hwd = 0 Then MsgBox "please open": Exit Sub
    dc = GetDC(hwd)
    
    Call GetWindowRect(hwd, Rct)
    
    If (Rct.Right - Rct.Left) <> 600 Then
       MsgBox "please open":
       hwd = -1: dc = -1
       Exit Sub
    End If
    
Debug.Print Rct.Bottom
Debug.Print Rct.Top
Debug.Print Rct.Left
End Sub

Private Sub cmd2_Click()
    Dim r&, k&
    Dim d As Integer
    Dim getRGB(0 To 799) As String

    ReDim ci(15) As Long

    Dim tp As Long
    'r = GetPixel(dc, 0, 0)
   'Debug.Print CStr(Hex(GetPixel(dc, 1, 1)))
   
  n = nameToHwndEx("大家来找茬", ci, 10)
  Debug.Print "number: " & CStr(n)
  
Do While n > -1
  If n <> 0 Then
     n = n - 1
     Call GetWindowRect(ci(n), Rct)
     If (Rct.Right - Rct.Left) = 800 Then Exit Do
  Else
    MsgBox "please open"
    Exit Sub
  End If
Loop
' top_hWnd ci(n)
Debug.Print ci(n)
step = getHWnd
hi = Rct.Bottom - Rct.Top
wi = Rct.Right - Rct.Left
Timer1.Interval = 20
'tp = (Rct.Top + Rct.Bottom) / 2
'
'dc = GetDC(hwd)
'For d = 0 To 799
'
'   Call SetMousePos(Rct.Left + d, tp)
'      mouse_event MOUSEEVENTF_LEFTDOWN, 0, 0, 0, 0
' Call waittime(tcik)
'      mouse_event MOUSEEVENTF_LEFTUP, 0, 0, 0, 0
'getRGB(d) = CStr(Hex(GetPixel(dc, d, tp)))
'
'
'Next d
'
'tp = tp + 1
'Call wr_txtEx(tit, "Anan", getRGB, 800)
    

    '        'Call SetMousePos(Rct.Right, (Rct.Bottom + Rct.Top) / 2)
    '        Call SetMousePos((Rct.Right + Rct.Left) / 2, Rct.Top - 1)
    '        mouse_event MOUSEEVENTF_LEFTDOWN, 0, 0, 0, 0
    '        Call waittime(tcik)
    '        mouse_event MOUSEEVENTF_LEFTUP, 0, 0, 0, 0
End Sub

Private Sub Form_Load()

    init_dir (tit)
    Me.BackColor = RGB(67, 66, 66)
    cmd1.BackColor = RGB(67, 66, 66)
    Me.Label1.BackColor = RGB(67, 66, 66)
     'setAttribute cmd1.hwnd, cmd1.BackColor, 0, 1
     setAttribute Me.hwnd, Me.BackColor, 256, 1
    
End Sub

Private Sub Form_Unload(Cancel As Integer)
    If step = getHWnd Then top_hWnd ci(n), False
End Sub

Private Sub Timer1_Timer()

 Call drawHwndLine(ci(n), s(0), 0, s(0), hi)
 Call drawHwndLine(ci(n), 0, s(1), wi, s(1))
  
End Sub

Private Sub Value1_valChange(ByVal v As Integer, ByVal id As Integer)
    s(id) = v
    If step = getHWnd Then
    
    
    End If
End Sub
Public Sub drawHwndLine(ByVal h As Long, x1%, y1%, x2%, y2%)
'
Dim p0 As POINTAPI
Dim dc As Long
    dc = GetDC(h)
   '
Call MoveToEx(dc, x1, y1, p0)
Call LineTo(dc, x2, y2)

End Sub


