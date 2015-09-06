VERSION 5.00
Begin VB.Form Form1 
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   780
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   1200
   LinkTopic       =   "Form1"
   Picture         =   "Form1.frx":0000
   ScaleHeight     =   780
   ScaleWidth      =   1200
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  '窗口缺省
   Begin VB.Timer fileTimer 
      Left            =   840
      Top             =   360
   End
   Begin VB.Label l1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "一个热吧"
      ForeColor       =   &H00000000&
      Height          =   180
      Index           =   0
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   720
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Const appName = "remindShow"
Dim con() As String
Dim ptcon() As String
Dim num As Integer

Dim t As Integer
                '''数组控件用object类型表示就好了
Private Function sizer(cs As Object, n As Integer)
'Me.Width = l1.Left + l1.Width
'Me.Height = l1.Height + l1.top
Const margin = 50

Dim h As Integer
h = 0 + margin
Dim dh As Integer
dh = l1(0).Height
Dim w As Integer
w = 0
Dim tw As Integer

    Dim i As Integer
      For i = 1 To n
          Load l1(i) '
          l1(i).top = h: h = h + dh
          l1(i).Left = 0 + margin '
          l1(i).Visible = True 'Load法不管原控件visible如何都是为False
          l1(i).Caption = con(i)
          
          tw = l1(i).Width
          If (tw > w) Then w = tw
      Next i

h = h + margin
w = w + margin * 2
'Me.Width = 2000 'w
'Me.Height = 1000 'h
Me.Width = w 'w
Me.Height = h 'h
Debug.Print Me.Width
Debug.Print Me.Height


End Function

Private Function writePt()
        ReDim ptcon(1)
        ptcon(0) = CStr(Me.Left)
        ptcon(1) = CStr(Me.top)
        Call wr_txtEx(appName, "pt", ptcon, 2)
End Function
Private Sub fileTimer_Timer()
    t = t + 1
    'If (t > 1) Then
    If (t > 30) Then
        t = 0
        writePt
    End If
    
End Sub

Private Sub Form_Load()
ReDim con(1)
con(0) = "1"
con(1) = "test"
Call init_txtEx(appName, "show", con, 2)
num = Val(con(0))
ReDim con(num)
Call init_txtEx(appName, "show", con, num + 1)


ReDim ptcon(1)
Call init_txtEx(appName, "pt", ptcon, 2)
Const offset = 10
Dim X As Integer, Y As Integer
X = Val(ptcon(0))
Y = Val(ptcon(1))
If X > (Screen.Width - offset) Then
    Me.Left = Screen.Width - offset
Else
    Me.Left = X
End If
If Y > (Screen.Height - offset) Then
    Me.top = Screen.Height - offset
Else
    Me.top = Y
End If


t = 0
fileTimer.Interval = 1000


Call sizer(l1, num)
Call top_Form(Me)
l1(0).Visible = False

Call setStartUp(True, appName)
End Sub

Private Sub Form_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
DoDrag Me
End Sub

Private Sub l1_MouseDown(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
DoDrag Me
End Sub

