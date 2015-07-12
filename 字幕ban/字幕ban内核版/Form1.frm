VERSION 5.00
Begin VB.Form Form1 
   BackColor       =   &H00C0FFC0&
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   3030
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   4560
   LinkTopic       =   "Form1"
   ScaleHeight     =   202
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   304
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  '´°¿ÚÈ±Ê¡
   Begin VB.Timer Timer2 
      Left            =   1080
      Top             =   2280
   End
   Begin VB.Timer Timer1 
      Interval        =   100
      Left            =   600
      Top             =   2160
   End
   Begin VB.HScrollBar HScroll2 
      Height          =   255
      Left            =   0
      TabIndex        =   1
      Top             =   240
      Value           =   1000
      Width           =   2415
   End
   Begin VB.HScrollBar HScroll1 
      Height          =   255
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Value           =   1000
      Width           =   2415
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Declare Function SetWindowPos Lib "user32" (ByVal hwnd As Long, ByVal hWndInsertAfter As Long, ByVal x As Long, ByVal y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long) As Long
Private Declare Function GetCursorPos Lib "user32" (lpPoint As POINTAPI) As Long
Private Type POINTAPI
x As Long
y As Long
End Type
Dim p As POINTAPI
Dim p2 As POINTAPI
Dim t As Integer

Private Sub Form_DblClick()
If HScroll1.Visible = False Then
HScroll1.Visible = True
HScroll2.Visible = True
Else: HScroll1.Visible = False
HScroll2.Visible = False
End If
t = 0
End Sub

Private Sub Form_Load()
t = 0
Timer2.Interval = 10
HScroll1.Max = Screen.Width
HScroll1.Min = Screen.Width / 12
HScroll1.Value = HScroll1.Min
HScroll2.Max = Screen.Height
HScroll2.Min = Screen.Height / 12
HScroll2.Value = HScroll2.Min
End Sub

Private Sub Form_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
t = 1
p2.x = x '/ Screen.TwipsPerPixelX
p2.y = y '/ Screen.TwipsPerPixelY
Timer2.Enabled = True
End Sub

Private Sub Form_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
t = 0
Timer2.Enabled = False
End Sub

Private Sub HScroll1_Change()
Me.Width = HScroll1.Value

End Sub
Private Sub HScroll2_Change()

Me.Height = HScroll2.Value
End Sub

Private Sub Timer1_Timer()
SetWindowPos Form1.hwnd, -1, 0, 0, 0, 0, 3
End Sub

Private Sub Timer2_Timer()
If t = 1 Then GetCursorPos p
Me.Left = (p.x - p2.x) * Screen.TwipsPerPixelX
Me.Top = (p.y - p2.y) * Screen.TwipsPerPixelY
End Sub
