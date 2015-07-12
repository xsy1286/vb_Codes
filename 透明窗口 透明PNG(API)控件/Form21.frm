VERSION 5.00
Begin VB.Form Form1 
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   6000
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   8955
   LinkTopic       =   "Form1"
   ScaleHeight     =   6000
   ScaleWidth      =   8955
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  '窗口缺省
   Begin VB.CommandButton Command1 
      Caption         =   "Command1"
      Height          =   735
      Left            =   720
      TabIndex        =   1
      Top             =   1920
      Width           =   2175
   End
   Begin 工程1.UserControl1 U1 
      Height          =   4575
      Left            =   1440
      TabIndex        =   0
      Top             =   720
      Width           =   6735
      _ExtentX        =   6376
      _ExtentY        =   5318
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Declare Function GetWindowLong Lib "user32" Alias "GetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long) As Long
Private Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
Private Declare Function SetLayeredWindowAttributes Lib "user32" (ByVal hwnd As Long, ByVal crKey As Long, ByVal bAlpha As Byte, ByVal dwFlags As Long) As Long
Const WS_EX_LAYERED = &H80000
Const GWL_EXSTYLE = (-20)
Const LWA_COLORKEY = &H1

Private Sub Form_Load()
Dim transcolor As Long
 transcolor = RGB(67, 66, 66) '必须与SetLayeredWindowAttributes第二个参数一致
 Me.BackColor = transcolor '
   Dim rtn As Long
   rtn = GetWindowLong(hwnd, GWL_EXSTYLE)
   rtn = rtn Or WS_EX_LAYERED
   SetWindowLong hwnd, GWL_EXSTYLE, rtn
   SetLayeredWindowAttributes hwnd, transcolor, 0, LWA_COLORKEY

U1.backsty = 1
Dim r As Long
r = U1.bc(50, 50, 100)
U1.url = "H:\123123.png"

End Sub



Private Sub U1_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
DoDrag Me
End Sub
