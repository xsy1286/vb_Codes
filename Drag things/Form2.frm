VERSION 5.00
Begin VB.Form Form2 
   BackColor       =   &H00000000&
   Caption         =   "Form2"
   ClientHeight    =   6360
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   11460
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   134
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form2"
   ScaleHeight     =   6360
   ScaleWidth      =   11460
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  '´°¿ÚÈ±Ê¡
   Visible         =   0   'False
   Begin VB.CommandButton cmdCommand1 
      Caption         =   "Command1"
      Height          =   360
      Left            =   -480
      TabIndex        =   3
      Top             =   480
      Width           =   990
   End
   Begin VB.VScrollBar Vrl1 
      Height          =   1455
      Left            =   360
      Max             =   100
      Min             =   1
      TabIndex        =   2
      Top             =   1200
      Value           =   10
      Width           =   255
   End
   Begin VB.PictureBox Pic1 
      AutoSize        =   -1  'True
      Height          =   3225
      Left            =   3000
      Picture         =   "Form2.frx":0000
      ScaleHeight     =   3165
      ScaleWidth      =   5625
      TabIndex        =   0
      TabStop         =   0   'False
      Top             =   1200
      Visible         =   0   'False
      Width           =   5685
   End
   Begin VB.PictureBox Pic2 
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      Height          =   5535
      Left            =   6240
      ScaleHeight     =   5535
      ScaleWidth      =   5775
      TabIndex        =   1
      TabStop         =   0   'False
      Top             =   840
      Width           =   5775
   End
End
Attribute VB_Name = "Form2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim a As Image
Dim d As Double

Private Sub cmdCommand1_Click()
Debug.Print CStr(cmdCommand1.Left)
cmdCommand1.Left = -100
End Sub

Private Sub Form_Click()
'Pic1.Width = Picture.Width * d
'Pic1.Height = Picture.Height * d
'Pic1.PaintPicture Picture, 0, 0, Picture.Width * d, Picture.Height * d
'Pic1.PaintPicture Picture, 0, 0, Picture.Width * d, Picture.Height * d
End Sub

Private Sub Form_Load()
'Picture = LoadPicture("C:\Users\Administrator\Desktop\2.bmp")
'Me.Picture = LoadPicture("C:\Users\Administrator\Desktop\2.bmp")


Me.Show
d = 1
Call zom
End Sub

Private Sub Vrl1_Change()
d = Vrl1.Value / 10
Call zom

End Sub

Private Sub zom()
Pic2.Left = (Me.Width - Pic1.Picture.Width * d) / 2
'Pic2.Top = (Me.Height - Pic1.Picture.Height * d) / 2
Pic2.Top = 0
Pic2.Width = Pic1.Picture.Width * d
Pic2.Height = Pic1.Picture.Width * d
Pic2.Refresh
Pic2.PaintPicture Pic1.Picture, 0, 0, Pic1.Picture.Width * d, Pic1.Picture.Width * d
End Sub
