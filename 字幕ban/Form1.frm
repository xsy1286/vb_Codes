VERSION 5.00
Begin VB.Form Form1 
   AutoRedraw      =   -1  'True
   BackColor       =   &H0080FF80&
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   4500
   ClientLeft      =   105
   ClientTop       =   105
   ClientWidth     =   7920
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   ScaleHeight     =   4500
   ScaleWidth      =   7920
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  '窗口缺省
   Begin VB.Timer Timer3 
      Left            =   2040
      Top             =   960
   End
   Begin VB.Timer Timer2 
      Left            =   2760
      Top             =   840
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Type POINTAPI
    X As Long
    Y As Long
End Type
Private Declare Function SetWindowPos Lib "user32" (ByVal hwnd As Long, ByVal hWndInsertAfter As Long, ByVal X As Long, ByVal Y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long) As Long
Const HWND_TOPMOST = -1
Const HWND_NOTOPMOST = -2
Const SWP_NOSIZE = &H1
Const SWP_NOMOVE = &H2
Const SWP_NOACTIVATE = &H10
Const HWND_TOP = 0
Const SWP_NOREDRAW = &H8
Const SWP_NOREPOSITION = &H200
Const SWP_NOZORDER = &H4
Private Declare Function GetCursorPos Lib "user32" (lpoint As POINTAPI) As Long

'Dim pa As POINTAPI '！！！定义结构必须分开定义
'Dim pv As POINTAPI
Dim pm1 As POINTAPI
Dim pm2 As POINTAPI
Dim k As Long
Dim hideflag As Boolean
Private Sub Form_Load()
 top_Form Me, False
End Sub

Private Sub Form_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
 frm1alltop = 1
If Button = vbLeftButton And movable = True Then
DoDrag Me
'MsgBox "movable=1"
ElseIf Button = vbRightButton Then

End If

End Sub
 
Private Sub Form_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
If Button = vbRightButton Then
'mymenu = Me.m1name
  Timer2.Interval = 0: frm1alltop = 0
  Timer3.Interval = 0
top_Form Me, False
 PopupMenu rBtuM.m1name
 
End If

End Sub



Private Sub Timer2_Timer()
If hideflag = False And frm1alltop = 1 Then
top_Form Me, True
'Debug.Print "1"
End If
End Sub

Private Sub Timer3_Timer() '单位100毫秒

If frs = 1 Then
Call GetCursorPos(pm1)
pm2 = pm1
frs = 0
k = p
Else
Call GetCursorPos(pm1)
'Ldbg pm1.X: Ldbg pm1.Y
End If

If hideflag = False Then
If ((pm1.X - pm2.X) ^ 2 + (pm1.Y - pm2.Y) ^ 2) > 4 Then
hideflag = True: Me.Hide
End If
ElseIf hideflag = True Then
If k = 0 Then
hideflag = False: Me.Show
frs = 1
Else
k = k - 1
End If

End If

pm2.X = pm1.X
pm2.Y = pm1.Y
End Sub



