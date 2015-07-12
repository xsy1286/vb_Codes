VERSION 5.00
Begin VB.UserControl Value 
   BackStyle       =   0  '透明
   ClientHeight    =   2595
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   3795
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   ScaleHeight     =   2595
   ScaleWidth      =   3795
   Begin VB.TextBox txt 
      Height          =   285
      Index           =   1
      Left            =   1320
      TabIndex        =   3
      Top             =   1350
      Width           =   1815
   End
   Begin VB.TextBox txt 
      Height          =   285
      Index           =   0
      Left            =   1320
      TabIndex        =   2
      Top             =   240
      Width           =   1815
   End
   Begin VB.HScrollBar h1 
      Height          =   255
      Index           =   1
      Left            =   1320
      TabIndex        =   1
      Top             =   1920
      Width           =   1815
   End
   Begin VB.HScrollBar h1 
      Height          =   255
      Index           =   0
      Left            =   1320
      TabIndex        =   0
      Top             =   810
      Width           =   1815
   End
   Begin VB.Label lb 
      Caption         =   "Label1"
      Height          =   195
      Index           =   1
      Left            =   240
      TabIndex        =   5
      Top             =   1440
      Width           =   465
   End
   Begin VB.Label lb 
      Caption         =   "Label1"
      Height          =   195
      Index           =   0
      Left            =   240
      TabIndex        =   4
      Top             =   240
      Width           =   465
   End
End
Attribute VB_Name = "Value"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
'#Const m_debug = "none"
Const k As Integer = &H2
Const hmin = 0
Const hmax = 1000

Public Event valChange(ByVal v As Integer, ByVal id As Integer)
Dim id As Integer


Private Sub h1_Change(Index As Integer)
txt(Index).Text = CStr(h1(Index).Value)
 RaiseEvent valChange(h1(Index).Value, Index)
End Sub

Private Sub txt_Change(Index As Integer)
h1(Index).Value = Val(txt(Index).Text)
  RaiseEvent valChange(h1(Index).Value, Index)
End Sub

Private Sub UserControl_Initialize()
Dim i As Integer
For i = 0 To k - 1
    lb(i) = "label" & CStr(i)
    txt(i) = CStr(i)
    With h1(i)
       .min = hmin
       .max = hmax
       .Value = i
    End With
Next
id = 0
End Sub
Public Property Get num() As Integer
num = id
End Property
Public Property Let num(ByVal a As Integer)
    If a < 0 Or a > (k - 1) Then MsgBox "Wrong Number": Exit Property
    id = a
    Me.hvalue = h1(id).Value '可如此非在某Property改变的其Property
End Property
Public Property Get hvalue() As Integer
    hvalue = h1(id).Value
End Property
Public Property Let hvalue(ByVal a As Integer)
  txt(id).Text = CStr(a)
  h1(id).Value = a
End Property

