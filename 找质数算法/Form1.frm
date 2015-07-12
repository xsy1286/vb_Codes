VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   5130
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   12330
   LinkTopic       =   "Form1"
   ScaleHeight     =   5130
   ScaleWidth      =   12330
   StartUpPosition =   3  '´°¿ÚÈ±Ê¡
   Begin VB.PictureBox Picture1 
      Height          =   4335
      Left            =   5760
      ScaleHeight     =   1.00000e5
      ScaleMode       =   0  'User
      ScaleWidth      =   10000
      TabIndex        =   4
      Top             =   240
      Width           =   6255
   End
   Begin VB.TextBox Text2 
      Height          =   495
      Left            =   3720
      TabIndex        =   3
      Top             =   1920
      Width           =   1455
   End
   Begin VB.TextBox Text1 
      Height          =   495
      Left            =   3720
      TabIndex        =   2
      Top             =   480
      Width           =   1455
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Command1"
      Height          =   495
      Left            =   3960
      TabIndex        =   1
      Top             =   3120
      Width           =   1215
   End
   Begin VB.ListBox List1 
      Height          =   3630
      Left            =   360
      Style           =   1  'Checkbox
      TabIndex        =   0
      Top             =   360
      Width           =   2775
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim a, b As Double
Private i, k As Double
Private j As Double
Dim bo As Boolean

Private Sub Command1_Click()
List1.Clear
k = 0
a = Val(Text1.Text): b = Val(Text2.Text)

If a > b Then a = a + b: b = a - b: a = a - b

If a <= 2 Then
List1.AddItem "2"
a = 3
End If

If a Mod 2 = 0 Then a = a + 1


For i = a To b Step 2

bo = False
For j = 3 To (Sqr(i)) Step 2
If i Mod j = 0 Then bo = True: Exit For
Next j
If bo = False Then k = k + 1: List1.AddItem CStr(i): Picture1.PSet (k, i), vbBlue


Next i
End Sub
