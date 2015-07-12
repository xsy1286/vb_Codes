VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   5355
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   6990
   LinkTopic       =   "Form1"
   ScaleHeight     =   5355
   ScaleWidth      =   6990
   StartUpPosition =   3  '´°¿ÚÈ±Ê¡
   Begin VB.Timer Timer1 
      Left            =   5640
      Top             =   1560
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Command2"
      Height          =   1215
      Left            =   -360
      TabIndex        =   3
      Top             =   -360
      Width           =   1455
   End
   Begin VB.HScrollBar H1 
      Height          =   615
      Left            =   1200
      TabIndex        =   2
      Top             =   3840
      Width           =   3855
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Command1"
      Height          =   1335
      Left            =   960
      TabIndex        =   0
      Top             =   600
      Width           =   2295
   End
   Begin VB.OLE OLE1 
      Height          =   975
      Left            =   3120
      TabIndex        =   1
      Top             =   2040
      Width           =   735
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim x As Long, y As Long
Dim t As Long
Dim serh() As Long
Private Type POINTAPI
 x As Long
 y As Long
End Type

Private Sub Command1_Click()
x = x + 1: y = y + 1
Debug.Print "x:" & CStr(x) & "  y:" & CStr(y)
sendMouse Command2.hwnd, WM_LBUTTONDOWN, x, y
sendMouse Command2.hwnd, WM_LBUTTONUP, x, y
End Sub

Private Sub Command2_Click()
Me.PSet (1581, 600), vbRed
MsgBox "1'"
End Sub

Private Sub Form_Load()
x = 0
y = 0
Timer1.Interval = 1000
End Sub

Private Sub Form_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
PSet (x, y), vbRed
End Sub

Private Sub Timer1_Timer()
t = t + 1
Debug.Print "t:" & CStr(t) & "s"
ReDim serh(10) As Long
Dim i
For i = 0 To 10
    serh(i) = 0
Next
If t = 10 Then
   Call nameToHwndEx("ms", serh, 5)
   Debug.Print "first hwnd got is:" & CStr(serh(0))
   Debug.Print "second hwnd got is:" & CStr(serh(1))
   Debug.Print "second hwnd got is:" & CStr(serh(2))
ElseIf t = 11 Then
    tmpDraw (serh(0))
    tmpDraw (serh(1))
    tmpDraw (serh(2))
    
   'Call sendMouse(Me.hwnd, WM_CLOSE, 350, 350)
   Call sendMouse(Command1.hwnd, WM_LBUTTONDOWN, 50, 50)
   Call sendMouse(Command1.hwnd, WM_LBUTTONUP, 50, 50)

   Call sendMouse(serh(0), WM_LBUTTONDOWN, 50, 50)
   Call sendMouse(serh(0), WM_LBUTTONUP, 50, 50)
   
'    Call sendMouse(serh(1), WM_LBUTTONDOWN, 1250, 223)
'   Call sendMouse(serh(1), WM_LBUTTONUP, 1250, 223)
'
'      Call sendMouse(serh(2), WM_LBUTTONDOWN, 1250, 223)
'   Call sendMouse(serh(2), WM_LBUTTONUP, 1250, 223)
End If


End Sub

Private Function tmpDraw(hwd As Long)
    If hwd <> 0 Then
   
    Dim dc As Long
    dc = GetDC(hwd)
    Call MoveToEx(dc, 500, 0, tmpP)
    Call LineTo(dc, 500, 600)
    Call ReleaseDC(hwd, dc)
    End If
End Function
