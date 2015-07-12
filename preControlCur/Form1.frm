VERSION 5.00
Begin VB.Form Form1 
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   6120
   ClientLeft      =   105
   ClientTop       =   105
   ClientWidth     =   9315
   LinkTopic       =   "Form1"
   ScaleHeight     =   6120
   ScaleWidth      =   9315
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  '窗口缺省
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim X As Long, Y As Long
Dim wm As Long

Dim WithEvents oControl As Timer
Attribute oControl.VB_VarHelpID = -1
Dim WithEvents t2 As Timer
Attribute t2.VB_VarHelpID = -1
Dim t As Long

Dim sDc As Long
Dim tmp As POINTAPI
Dim w As Long, h As Long
Dim wx As Integer, hy As Integer
Dim action As String

Private Sub Form_Click()
    frmScr.Show
End Sub

Private Sub Form_Load()
    ' MsgBox "start"
    Call toTray(Me, "screenDC", True)
    '    Load JustforMenu
    '    JustforMenu.Hide
    t = 0
    wm = 0
    w = Screen.Width
    wx = Screen.TwipsPerPixelX
    h = Screen.Height
    hy = Screen.TwipsPerPixelY
    
    
    Call hook  'start hook
    
    
          '                        vb-你控件的工程名    TextBox－你控件的类名     newText－生成后的控件名称
   Set oControl = Controls.Add("VB.timer", "t1", Me)
   Set t2 = Controls.Add("VB.timer", "t2name", Me)
   
   t2.Interval = 100
   
    With oControl
         .Interval = 10
    End With
    
    oControl.Enabled = True
    
    top_Form Line1: top_Form Line2: top_Form Line3: top_Form Line4:
    
    Line1.Height = hy
    Line2.Height = hy
    Line3.Width = wx
    Line4.Width = wx
    
    Line1.Show
     Line2.Show
      Line3.Show
       Line4.Show
       
    frmScr.Show
    
    Call GetCursorPos(tmp)
    X = tmp.X
    Y = tmp.Y
    wm = WM_MOUSEMOVE
    Call oControl_Timer
    
    t2.Enabled = True
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Dim Msg As Long
    Msg = X / 15
    Debug.Print "m"
    If Msg = WM_LBUTTONUP Then
        Call UnHooK
        End
    End If
'    If Msg = WM_RBUTTONDOWN Then     '有执行，但菜单弹出即消失
'      top_Form Line1, False: top_Form Line2, False: top_Form Line3, False: top_Form Line4, False:
'           PopupMenu JustforMenu.n_pop
'    End If
End Sub
Public Function fnHook(x1 As Long, y1 As Long, mes As Long) As Long
    X = x1
    Y = y1
    wm = mes
End Function

Public Sub try()

End Sub


Private Sub Form_Unload(Cancel As Integer)
 Call UnHooK  'improtant
End Sub

Private Sub lb1_Click()

End Sub

Private Sub oControl_Timer()
    top_Form Line1: top_Form Line2: top_Form Line3: top_Form Line4:
    Dim r As Long

    '    sDc = GetDC(0)
    '    r = MoveToEx(sDc, 0, y, tmp)
    '    r = LineTo(sDc, w, y)
    '    r = MoveToEx(sDc, x, 0, tmp)
    '    r = LineTo(sDc, x, h)
    '    r = ReleaseDC(0, sDc)

    '    frmScr.Line1.x1 = 0: frmScr.Line1.X2 = (X - 1) * wx  'this way cannot topmost
    '    frmScr.Line1.y1 = Y * hy: frmScr.Line1.Y2 = Y * hy
    '
    '    frmScr.Line2.x1 = X * wx: frmScr.Line2.X2 = X * wx
    '    frmScr.Line2.y1 = 0: frmScr.Line2.Y2 = (Y - 1) * hy
    '
    '    frmScr.Line3.x1 = (X + 1) * wx: frmScr.Line3.X2 = w
    '    frmScr.Line3.y1 = Y * hy: frmScr.Line3.Y2 = Y * hy
    '
    '    frmScr.Line4.x1 = X * wx: frmScr.Line4.X2 = X * wx
    '    frmScr.Line4.y1 = (Y + 1) * hy: frmScr.Line4.Y2 = h
    If wm = WM_LBUTTONDOWN Then
        action = "WM_LBUTTONDOWN"
    ElseIf wm = WM_LBUTTONUP Then
        action = "WM_LBUTTONUP"
    ElseIf wm = WM_LBUTTONDBLCLK Then
        action = "WM_LBUTTONDBLCLK"
    ElseIf wm = WM_RBUTTONDOWN Then
        action = "WM_RBUTTONDOWN"
    ElseIf wm = WM_RBUTTONUP Then
        action = "WM_RBUTTONUP"
    ElseIf wm = WM_RBUTTONDBLCLK Then
        action = "WM_RBUTTONDBLCLK"
    Else
        If X > 2 Then
            Line1.Left = 0:  Line1.Width = (X - 1) * wx
            Line1.top = hy * Y
            Line1.Show
        Else
            Line1.Hide
        End If
        Line2.Left = (X + 1) * wx: Line2.Width = w
        Line2.top = hy * Y

        If Y > 2 Then
            Line3.top = 0:  Line3.Height = (Y - 1) * hy
            Line3.Show
            Line3.Left = wx * X
        Else
            Line3.Hide
        End If
        Line4.top = (Y + 1) * hy:  Line4.Height = h
        Line4.Left = wx * X

        frmScr.lb1.Caption = "x:" & CStr(X) & " y:" & CStr(Y)

        wm = 0
        Exit Sub
    End If

    Call txtAppend(main.appName, main.txtnam, CStr(t / 10) & "s  x:" & CStr(X) & " y:" & CStr(Y) & " action is:" & action)
    wm = 0
End Sub

Private Sub t2_Timer()
 
 t = t + 1
 
End Sub
