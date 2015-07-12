VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   7950
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   13050
   LinkTopic       =   "Form1"
   ScaleHeight     =   7950
   ScaleWidth      =   13050
   Begin VB.DirListBox Dir1 
      Height          =   1350
      Left            =   1320
      TabIndex        =   10
      Top             =   3600
      Width           =   2775
   End
   Begin VB.HScrollBar HScroll4 
      Height          =   495
      Left            =   11160
      Max             =   32760
      TabIndex        =   9
      Top             =   5520
      Value           =   10000
      Width           =   1575
   End
   Begin VB.HScrollBar HScroll3 
      Height          =   495
      Left            =   11040
      Max             =   32760
      TabIndex        =   8
      Top             =   4680
      Value           =   10000
      Width           =   1575
   End
   Begin VB.HScrollBar HScroll2 
      Height          =   375
      Left            =   960
      Max             =   1000
      TabIndex        =   7
      Top             =   6120
      Width           =   1935
   End
   Begin VB.HScrollBar HScroll1 
      Height          =   375
      Left            =   0
      Max             =   1000
      TabIndex        =   6
      Top             =   5520
      Width           =   1695
   End
   Begin VB.TextBox Text1 
      Height          =   855
      Left            =   240
      TabIndex        =   5
      Text            =   "Text1"
      Top             =   4320
      Width           =   2055
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Command3"
      Height          =   975
      Left            =   600
      TabIndex        =   4
      Top             =   2520
      Width           =   90
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Command2"
      Height          =   975
      Left            =   9000
      TabIndex        =   3
      Top             =   1080
      Width           =   2175
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Command1"
      Height          =   495
      Left            =   480
      TabIndex        =   2
      Top             =   480
      Width           =   975
   End
   Begin VB.PictureBox Picture2 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      Height          =   3135
      Left            =   2400
      ScaleHeight     =   3075
      ScaleWidth      =   5595
      TabIndex        =   1
      Top             =   480
      Width           =   5655
   End
   Begin VB.PictureBox Picture1 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      Height          =   3735
      Left            =   3600
      ScaleHeight     =   245
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   469
      TabIndex        =   0
      Top             =   4200
      Width           =   7095
   End
   Begin VB.Shape Shape1 
      BorderWidth     =   40
      Height          =   735
      Left            =   600
      Top             =   1680
      Width           =   1455
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Declare Function PrintWindow Lib "user32" (ByVal hwnd As Long, ByVal hdcBlt As Long, ByVal nFlags As Long) As Long
Private Declare Function BitBlt Lib "gdi32" (ByVal hDestDC As Long, ByVal X As Long, ByVal Y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal dwRop As Long) As Long
Private Declare Function StretchBlt Lib "gdi32" (ByVal hdc As Long, ByVal X As Long, ByVal Y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal nSrcWidth As Long, ByVal nSrcHeight As Long, ByVal dwRop As Long) As Long
Private Declare Function GetDC Lib "user32" (ByVal hwnd As Long) As Long
Private Declare Function GetCursorPos Lib "user32" (lpPonit As POINTAPI) As Long
Const Srccopy = &HCC0020
Private Type POINTAPI
X As Long: Y As Long
End Type
Dim x1, y1, x2, y2 As Single
Dim t As Integer

Private Sub Command3_Click()

Picture2.Width = Val(HScroll3.Value - HScroll1.Value)
Text1.Text = Str(HScroll3.Value - HScroll1.Value)
Picture2.Height = Val(HScroll4.Value - HScroll2.Value)
Me.Hide
BitBlt Picture2.hdc, (0 - HScroll1.Value), (0 - HScroll2.Value), Picture2.Width, Picture2.Height, GetDC(0), 0, 0, vbSrcCopy
SavePicture Picture2.Image, (Dir1.Path + "\a.bmp")
Me.Show
End Sub



Private Sub Form_Load()
Me.Width = Screen.Width
Me.Height = Screen.Height

Picture2.AutoRedraw = True
t = 0
End Sub

Private Sub Form_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
'If t = 0 Then t = 1:
x1 = X: y1 = Y
Text1.Text = Str(x1) + "  " + Str(y1)

End Sub

Private Sub Form_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
If t = 1 Then
t = 0: x2 = X: y2 = Y
BitBlt Picture2.hdc, 0, 0, x2, y2, GetDC(0), 0, 0, vbSrcCopy
SavePicture Picture2.Image, "c:\aaa.bmp"
End If
End Sub

Private Sub HScroll1_Change()
Picture2.Width = Val(HScroll3.Value - HScroll1.Value)
Text1.Text = Str(HScroll3.Value - HScroll1.Value)
End Sub
