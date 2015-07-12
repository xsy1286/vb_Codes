VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "中断初值计算"
   ClientHeight    =   3555
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   4560
   LinkTopic       =   "Form1"
   ScaleHeight     =   3555
   ScaleWidth      =   4560
   StartUpPosition =   3  '窗口缺省
   Begin VB.CommandButton Command2 
      Caption         =   "计算中断间隔"
      Height          =   495
      Left            =   2640
      TabIndex        =   9
      Top             =   2760
      Width           =   1455
   End
   Begin VB.CommandButton Command1 
      Caption         =   "计算THX,TLX"
      Height          =   495
      Left            =   360
      TabIndex        =   8
      Top             =   2760
      Width           =   1695
   End
   Begin VB.TextBox Text4 
      Height          =   375
      Left            =   2880
      TabIndex        =   7
      Text            =   "Text4"
      Top             =   2040
      Width           =   1095
   End
   Begin VB.TextBox Text3 
      Height          =   375
      Left            =   2880
      TabIndex        =   6
      Text            =   "Text3"
      Top             =   1440
      Width           =   1095
   End
   Begin VB.TextBox Text2 
      Height          =   375
      Left            =   2880
      TabIndex        =   5
      Text            =   "Text2"
      Top             =   840
      Width           =   1095
   End
   Begin VB.TextBox Text1 
      Height          =   375
      Left            =   2880
      TabIndex        =   3
      Text            =   "Text1"
      Top             =   240
      Width           =   1095
   End
   Begin VB.Label Label5 
      BackStyle       =   0  'Transparent
      Caption         =   "16位定时器/计数器"
      Height          =   255
      Left            =   3030
      TabIndex        =   10
      Top             =   3360
      Width           =   1935
   End
   Begin VB.Label Label4 
      Caption         =   "计时<65.5326ms"
      Height          =   375
      Left            =   480
      TabIndex        =   4
      Top             =   2040
      Width           =   1575
   End
   Begin VB.Label Label3 
      Caption         =   "T晶振（us)"
      Height          =   375
      Left            =   480
      TabIndex        =   2
      Top             =   1440
      Width           =   1095
   End
   Begin VB.Label Label2 
      Caption         =   "TLX"
      Height          =   375
      Left            =   480
      TabIndex        =   1
      Top             =   840
      Width           =   735
   End
   Begin VB.Label Label1 
      Caption         =   "THX"
      Height          =   255
      Left            =   480
      TabIndex        =   0
      Top             =   360
      Width           =   615
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
On Error Resume Next
Dim a, b, c, d As Double
a = Val(Text1.Text)
b = Val(Text2.Text)
c = Val(Text3.Text)
d = Val(Text4.Text)

a = (65536 - (d / c) * 1000) \ 256
b = (65536 - (d / c) * 1000) Mod 256

Text1.Text = CStr(a)
Text2.Text = CStr(b)

End Sub

Private Sub Command2_Click()
On Error Resume Next
Dim a, b, c, d As Double
a = Val(Text1.Text)
b = Val(Text2.Text)
c = Val(Text3.Text)
d = Val(Text4.Text)

d = (65536 - (a * 256 + b) * c) / 1000

Text4.Text = CStr(d)

End Sub
