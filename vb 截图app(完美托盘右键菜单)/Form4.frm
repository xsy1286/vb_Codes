VERSION 5.00
Begin VB.Form Form4 
   BorderStyle     =   0  'None
   Caption         =   "Form4"
   ClientHeight    =   1455
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   945
   LinkTopic       =   "Form4"
   ScaleHeight     =   1455
   ScaleWidth      =   945
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  '´°¿ÚÈ±Ê¡
   Begin VB.Timer Timer1 
      Left            =   600
      Top             =   720
   End
   Begin VB.CommandButton Command1 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFC0&
      Caption         =   "Command1"
      Height          =   495
      Index           =   0
      Left            =   0
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   0
      Width           =   1215
   End
End
Attribute VB_Name = "Form4"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Const inb = 3
Dim i As Integer
Dim pix As Integer
Private hook As Long
Dim ptf42 As POINTAPI

Private Sub Command1_Click(Index As Integer)
Select Case Index
Case 1
Unload Me
Form1.Show
Case 2
Unload Me
Case 3
End

End Select


End Sub

Private Sub Command1_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
Command1(pix).BackColor = &HFFFFC0
Command1(Index).BackColor = &HFFFF00
pix = Index

End Sub

Private Sub Form_Load()

Me.Top = ty * Screen.TwipsPerPixelY - Me.Height
Me.Left = tx * Screen.TwipsPerPixelX

Debug.Print Str(pt.Y)
Debug.Print Str(Me.Top)
Command1(0).Width = Me.Width
Command1(0).Visible = True

Dim hght As Integer
hght = Me.Height / inb

For i = 1 To inb
Load Command1(i)
With Command1(i)
.Top = (i - 1) * hght
.Visible = True 'visible default false
End With
Next i
Command1(0).Visible = False
pix = 0

'capation
Command1(1).Caption = "Show"
Command1(2).Caption = "Cancel"
Command1(3).Caption = "End"
Timer1.Interval = 20
Timer1.Enabled = True

Me.Show

 hook = SetWindowsHookEx(WH_MOUSE_LL, AddressOf MouseHookf4, App.hInstance, 0)
  
SetWindowPos Me.hwnd, HWND_TOPMOST, 0, 0, 0, 0, SWP_NOACTIVATE Or SWP_SHOWWINDOW Or SWP_NOMOVE Or SWP_NOSIZE
End Sub

Private Sub Form_Terminate()
  UnhookWindowsHookEx hook
End Sub

Private Sub Form_Unload(Cancel As Integer)

  UnhookWindowsHookEx hook
  Timer1.Enabled = False
End Sub

Private Sub Timer1_Timer()
GetCursorPos ptf42
If _
 ptf42.X > Form4.Left / Screen.TwipsPerPixelX And _
  ptf42.X < (Form4.Left + Form4.Width) / Screen.TwipsPerPixelX And _
  ptf42.Y > Form4.Top / Screen.TwipsPerPixelY And _
  ptf42.Y < (Form4.Top + Form4.Height) / Screen.TwipsPerPixelY _
Then
Else:

Command1(pix).BackColor = &HFFFFC0

End If
End Sub
