VERSION 5.00
Begin VB.Form Form1 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Form1"
   ClientHeight    =   765
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   2040
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   765
   ScaleWidth      =   2040
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  '´°¿ÚÈ±Ê¡
   Begin VB.CommandButton Command4 
      Caption         =   "tmp"
      Height          =   1335
      Left            =   720
      TabIndex        =   9
      Top             =   1080
      Width           =   495
   End
   Begin VB.TextBox Text1 
      Height          =   375
      Left            =   1440
      TabIndex        =   8
      Text            =   "67"
      Top             =   120
      Width           =   495
   End
   Begin VB.Timer Timer2 
      Left            =   2520
      Top             =   240
   End
   Begin VB.Timer Timer1 
      Left            =   120
      Top             =   120
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Start"
      Height          =   495
      Left            =   0
      TabIndex        =   7
      Top             =   120
      Width           =   1335
   End
   Begin VB.PictureBox Picdown 
      Height          =   855
      Left            =   4800
      ScaleHeight     =   795
      ScaleWidth      =   1275
      TabIndex        =   6
      Top             =   3840
      Width           =   1335
   End
   Begin VB.PictureBox Picup 
      Height          =   855
      Left            =   2640
      ScaleHeight     =   795
      ScaleWidth      =   1275
      TabIndex        =   5
      Top             =   3360
      Width           =   1335
   End
   Begin VB.PictureBox Pic 
      Height          =   975
      Left            =   1560
      ScaleHeight     =   915
      ScaleWidth      =   1395
      TabIndex        =   4
      Top             =   4200
      Width           =   1455
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Load"
      Height          =   855
      Left            =   360
      TabIndex        =   3
      Top             =   3120
      Width           =   375
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Command1"
      Height          =   255
      Left            =   600
      TabIndex        =   2
      Top             =   3600
      Width           =   1815
   End
   Begin VB.PictureBox Picture2 
      BorderStyle     =   0  'None
      Height          =   2775
      Left            =   5760
      ScaleHeight     =   2775
      ScaleWidth      =   2895
      TabIndex        =   1
      Top             =   720
      Width           =   2895
   End
   Begin VB.PictureBox Picture1 
      Height          =   2175
      Left            =   840
      ScaleHeight     =   2115
      ScaleWidth      =   3555
      TabIndex        =   0
      Top             =   2760
      Width           =   3615
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim i As Integer
Dim j As Integer
Private Sub Command1_Click()
Picture2.PaintPicture Picture1.Picture, 0, 0
Picture2.PaintPicture Picture1.Picture, 0, Picture1.Height
SavePicture Picture2.Image, "C:\Users\Administrator\Desktop\2.bmp"
End Sub

Private Sub Command2_Click()
Picture1.Picture = LoadPicture("C:\Users\Administrator\Desktop\1.bmp")
Picture2.Width = Picture1.ScaleWidth
Picture2.Height = Picture1.ScaleHeight * 2
End Sub

Private Sub Command3_Click()

If Command3.Caption = "Start" Then
Me.Caption = "downloading"
Command3.Caption = "downloading"
j = Val(Text1.Text)
Timer2.Interval = 300
End If

End Sub

Private Sub Command4_Click()
Call inpic
SavePicture Pic.Image, "C:\Users\Administrator\Desktop\2.bmp"
End Sub



Private Sub Form_Load()
mid_Form Me
Me.Caption = "Document Download"
i = 0


Timer1.Interval = 10

Picture2.AutoRedraw = True

Picture1.AutoRedraw = True
Pic.AutoRedraw = True
Picup.AutoRedraw = True
Picdown.AutoRedraw = True

Picture1.AutoSize = True
Pic.AutoSize = True
Picup.AutoSize = True
Picdown.AutoSize = True

Picture1.BorderStyle = 0

Picture2.BorderStyle = 0


Pic.Width = 860 * Screen.TwipsPerPixelX
Pic.Height = 1362 * Screen.TwipsPerPixelY
End Sub
Private Sub getup(a As Long, b As Long, c As Long, d As Long)
BitBlt Picup.hdc, a, b, c, d, GetDC(0), 0, 0, vbSrcCopy
End Sub
Private Sub getdown(a As Long, b As Long, c As Long, d As Long)
BitBlt Picdown.hdc, a, b, c, d, GetDC(0), 0, 0, vbSrcCopy
End Sub

Private Sub Timer1_Timer()
top_Form Me
End Sub

Private Sub Timer2_Timer()
i = i + 1
If i <= j Then
Call inpic
SavePicture Pic.Image, "D:\2\" & CStr(i) & ".bmp"

SetCursorPos 800, 880
Call waittime(0.1)
VirtualClickMouse MOUSEEVENTF_LEFTDOWN, MOUSEEVENTF_LEFTUP
Else
Me.Caption = "Finish"
Command3.Caption = "Start"
i = 0
Timer2.Interval = 0
End If
End Sub
Private Sub inpic()
BitBlt Pic.hdc, -20, -128, 860 * Screen.TwipsPerPixelX, 1362 * Screen.TwipsPerPixelY, GetDC(0), 0, 0, vbSrcCopy
End Sub
