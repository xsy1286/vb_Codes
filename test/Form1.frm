VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   1260
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   2415
   LinkTopic       =   "Form1"
   ScaleHeight     =   1260
   ScaleWidth      =   2415
   StartUpPosition =   3  '´°¿ÚÈ±Ê¡
   Begin VB.Timer Timer1 
      Left            =   480
      Top             =   960
   End
   Begin VB.Label Label1 
      Caption         =   "Label1"
      Height          =   495
      Left            =   240
      TabIndex        =   0
      Top             =   360
      Width           =   1455
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim p As POINTAPI

Private Sub Form_Load()
Timer1.Interval = 20
top_Form Me
End Sub

Private Sub Form_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
Dim l As Long
l = GetCursorPos(p)
Print CStr(p.x) & "   " & CStr(p.y)
End Sub

Private Sub Timer1_Timer()
top_Form Me
Dim l As Long
l = GetCursorPos(p)
Label1.Caption = CStr(p.x) & "   " & CStr(p.y)
End Sub
