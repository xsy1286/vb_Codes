VERSION 5.00
Begin VB.Form Form2 
   Caption         =   "Color"
   ClientHeight    =   2625
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   2775
   LinkTopic       =   "Form2"
   ScaleHeight     =   2625
   ScaleWidth      =   2775
   StartUpPosition =   3  '窗口缺省
   Begin VB.CommandButton Command2 
      Caption         =   "取消"
      Height          =   495
      Left            =   1680
      TabIndex        =   4
      Top             =   1800
      Width           =   735
   End
   Begin VB.CommandButton Command1 
      Caption         =   "确定"
      Height          =   495
      Left            =   360
      TabIndex        =   3
      Top             =   1800
      Width           =   735
   End
   Begin VB.HScrollBar HScroll3 
      Height          =   255
      Left            =   480
      TabIndex        =   2
      Top             =   1200
      Width           =   1695
   End
   Begin VB.HScrollBar HScroll2 
      Height          =   255
      Left            =   480
      TabIndex        =   1
      Top             =   720
      Width           =   1695
   End
   Begin VB.HScrollBar HScroll1 
      Height          =   255
      Left            =   480
      TabIndex        =   0
      Top             =   240
      Width           =   1695
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "R       G       B"
      Height          =   1335
      Left            =   240
      TabIndex        =   5
      Top             =   240
      Width           =   255
   End
End
Attribute VB_Name = "Form2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim i As Integer
Private Sub Command1_Click()
Dim re As Long
r = HScroll1.Value

g = HScroll2.Value

b = HScroll3.Value

re = txtPrint("Tab", "ColorR", CStr(r))
re = txtPrint("Tab", "ColorG", CStr(g))
re = txtPrint("Tab", "ColorB", CStr(b))

For i = 0 To 2
Form1.Text1(i).ForeColor = RGB(r, g, b)
Next i

Unload Me
End Sub

Private Sub Command2_Click()
HScroll1.Value = r

HScroll2.Value = g

HScroll3.Value = b

Unload Me
End Sub

Private Sub Form_Load()
Me.Left = Form1.Left + 1320
Me.Top = Form1.Top + 1200

HScroll1.Min = 0: HScroll1.Max = 255
HScroll2.Min = 0: HScroll2.Max = 255
HScroll3.Min = 0: HScroll3.Max = 255

HScroll1.Value = r

HScroll2.Value = g

HScroll3.Value = b

End Sub


