VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   3090
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   ScaleHeight     =   3090
   ScaleWidth      =   4680
   StartUpPosition =   3  '´°¿ÚÈ±Ê¡
   Begin VB.Timer Timer1 
      Left            =   2760
      Top             =   960
   End
   Begin VB.TextBox Text1 
      Height          =   495
      Left            =   720
      TabIndex        =   1
      Text            =   "Text1"
      Top             =   480
      Width           =   1335
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Command1"
      Height          =   975
      Left            =   1080
      TabIndex        =   0
      Top             =   1440
      Width           =   1455
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
  
  

Private Sub Form_Load()
OldProc = GetWindowLong(Me.hwnd, GWL_WNDPROC)
Dim r As Long
r = SetWindowLong(Me.hwnd, GWL_WNDPROC, AddressOf WndProc)
Timer1.Interval = 1000
End Sub

Private Sub Form_Unload(Cancel As Integer)
Dim r As Long
r = SetWindowLong(Me.hwnd, GWL_WNDPROC, OldProc)

End Sub

