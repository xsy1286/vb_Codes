VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   6765
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   11055
   LinkTopic       =   "Form1"
   OLEDropMode     =   1  'Manual
   ScaleHeight     =   6765
   ScaleWidth      =   11055
   StartUpPosition =   3  '窗口缺省
   Begin VB.TextBox txtText1 
      BackColor       =   &H0000FFFF&
      Height          =   615
      Left            =   5880
      TabIndex        =   6
      Text            =   "Text1"
      Top             =   2640
      Width           =   1455
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Command2"
      Height          =   1095
      Left            =   840
      TabIndex        =   5
      Top             =   3000
      Width           =   2175
   End
   Begin VB.CheckBox Check1 
      Caption         =   "Check1"
      DragMode        =   1  'Automatic
      Height          =   615
      Left            =   5160
      TabIndex        =   4
      Top             =   840
      Width           =   1695
   End
   Begin VB.PictureBox Picture3 
      DragMode        =   1  'Automatic
      Height          =   1215
      Left            =   2760
      ScaleHeight     =   1155
      ScaleWidth      =   2235
      TabIndex        =   3
      Top             =   4680
      Width           =   2295
   End
   Begin VB.PictureBox Picture2 
      DragMode        =   1  'Automatic
      Height          =   1455
      Left            =   2880
      ScaleHeight     =   1395
      ScaleWidth      =   1635
      TabIndex        =   2
      Top             =   2760
      Width           =   1695
   End
   Begin VB.PictureBox Picture1 
      DragMode        =   1  'Automatic
      Height          =   1815
      Left            =   2880
      Picture         =   "Form1.frx":0000
      ScaleHeight     =   1755
      ScaleWidth      =   1395
      TabIndex        =   1
      Top             =   360
      Width           =   1455
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Command1"
      Height          =   975
      Left            =   720
      TabIndex        =   0
      Top             =   1080
      Width           =   1815
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim cka As CheckBox

Private Sub Command1_Click()
Dim f As Form1
Set f = New Form1
f.Show

End Sub

Private Function ckf()
L: Print "line"
End Function

Private Sub Command2_Click()
Form2.Show
End Sub

Private Sub Form_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)

cka.Left = X
cka.Top = Y + cka.Height
End Sub

Private Sub Picture3_DragDrop(Source As Control, X As Single, Y As Single)
If TypeOf Source Is PictureBox Then      ' 将 Picture3 位图设置为与源控件相同。
Picture3.Picture = Source.Picture
End If
End Sub

Private Sub Form_DragDrop(Source As Control, X As Single, Y As Single)
Dim ck As CheckBox
Set ck = Source
Set cka = Source
ck.Left = X
ck.Top = Y
End Sub
Private Sub Form_OLEDragDrop(Data As DataObject, Effect As Long, Button As Integer, Shift As Integer, X As Single, Y As Single)
Dim i As Long
On Error GoTo Errlog
For i = 1 To Data.Files.Count
Print Data.Files(i)
Next
Call ckf
Exit Sub
Errlog:
Call whenErr(Err.Number, "draging")
End Sub

