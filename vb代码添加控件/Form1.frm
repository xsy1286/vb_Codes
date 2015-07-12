VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   3030
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   4560
   LinkTopic       =   "Form1"
   ScaleHeight     =   3030
   ScaleWidth      =   4560
   StartUpPosition =   3  '´°¿ÚÈ±Ê¡
   Begin VB.CommandButton cmdCommand1 
      Caption         =   "Command1"
      Height          =   360
      Left            =   2760
      TabIndex        =   1
      Top             =   1080
      Width           =   990
   End
   Begin VB.TextBox Text1 
      DragMode        =   1  'Automatic
      Height          =   375
      Index           =   0
      Left            =   720
      TabIndex        =   0
      Text            =   "Text1"
      Top             =   720
      Width           =   975
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub cmdCommand1_Click()
Static i
i = i + 1
Load Text1(i)
Text1(0).Left = 0
'With Text1(i)
''.Top = 100 + 100 * i
''.Left = 100#
''.Visible = True
'End With


End Sub

Private Sub Form_Load()
'Text1.Index = 0


End Sub
