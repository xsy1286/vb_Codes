VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   3030
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   4560
   LinkTopic       =   "Form1"
   ScaleHeight     =   3030
   ScaleWidth      =   4560
   StartUpPosition =   3  '窗口缺省
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


'当双击托盘时恢复原状
Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

    Dim msg As Long
    msg = X / 15
    If msg = WM_LBUTTONDBLCLK Then
     Debug.Print "1"
    Me.WindowState = 0
    Me.Show
    Shell_NotifyIcon NIM_DELETE, p_Tray
    End If
End Sub
'当点最小化是为托盘状态
Private Sub Form_Resize()
    Call toTray(Me, "MyApp")
End Sub

