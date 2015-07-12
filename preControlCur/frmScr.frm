VERSION 5.00
Begin VB.Form frmScr 
   BorderStyle     =   0  'None
   Caption         =   "Form2"
   ClientHeight    =   3030
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   4560
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form2"
   ScaleHeight     =   3030
   ScaleWidth      =   4560
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  '´°¿ÚÈ±Ê¡
   Begin VB.Line Line4 
      X1              =   1920
      X2              =   2400
      Y1              =   2640
      Y2              =   1440
   End
   Begin VB.Line Line3 
      X1              =   2520
      X2              =   3720
      Y1              =   2400
      Y2              =   1800
   End
   Begin VB.Label lb1 
      BackStyle       =   0  'Transparent
      Caption         =   "×ø±ê"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   10.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Left            =   1560
      TabIndex        =   0
      Top             =   840
      Width           =   2220
   End
   Begin VB.Line Line2 
      X1              =   840
      X2              =   3360
      Y1              =   2160
      Y2              =   1920
   End
   Begin VB.Line Line1 
      X1              =   960
      X2              =   960
      Y1              =   1080
      Y2              =   2760
   End
End
Attribute VB_Name = "frmScr"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim bcl As Long

Private Sub Form_Load()
    bcl = RGB(100, 100, 100)
    'Line1.BorderColor = bcl
    Me.BackColor = bcl
   ' Call all_Screen(frmScr)
    'Me.WindowState = 2
    Call top_hWnd(Me.hwnd, True)
    
   ' Call setAttribute(Me.hwnd, bcl, 70, 1)
    
    Me.Left = Screen.Width * 7 / 8
    Me.top = Screen.Height - lb1.Height
    Me.Width = lb1.Width
    Me.Height = 600
    lb1.Left = 0
    lb1.top = 0
    
    Me.lb1.AutoSize = False
End Sub

