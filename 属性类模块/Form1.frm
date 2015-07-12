VERSION 5.00
Begin VB.Form Form1 
   BackColor       =   &H00FFFF80&
   Caption         =   "Form1"
   ClientHeight    =   5565
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   6960
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   ScaleHeight     =   5565
   ScaleWidth      =   6960
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  '窗口缺省
   Begin VB.HScrollBar h 
      Height          =   375
      Left            =   1440
      TabIndex        =   7
      Top             =   4800
      Width           =   1455
   End
   Begin 工程1.HValue H1 
      Height          =   1095
      Left            =   360
      TabIndex        =   6
      Top             =   3000
      Width           =   3015
      _extentx        =   5106
      _extenty        =   1931
   End
   Begin VB.CommandButton Command1 
      Caption         =   "&Command1"
      Height          =   375
      Left            =   1800
      TabIndex        =   3
      Top             =   120
      Width           =   1095
   End
   Begin VB.Frame Frame1 
      Caption         =   "Frame1"
      Height          =   1815
      Left            =   1680
      TabIndex        =   1
      Top             =   720
      Width           =   1935
      Begin VB.DriveListBox Drive1 
         Height          =   315
         Left            =   720
         TabIndex        =   5
         Top             =   120
         Width           =   750
      End
      Begin VB.CheckBox chkCheck1 
         Caption         =   "Check1"
         Height          =   255
         Left            =   840
         TabIndex        =   4
         Top             =   1320
         Width           =   1335
      End
      Begin VB.CommandButton cmdCommand1 
         Caption         =   "&Command1"
         Height          =   375
         Left            =   360
         TabIndex        =   2
         Top             =   480
         Width           =   1095
      End
   End
   Begin VB.CommandButton cmd 
      Caption         =   "1"
      Height          =   360
      Left            =   360
      TabIndex        =   0
      Top             =   1080
      Width           =   990
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim c
Private att As New valueCom
Attribute att.VB_VarHelpID = -1
Private Sub cmd_Click()
Call att.addValCon(1, 1, 1, 1)

End Sub

Private Sub Form_Load()
Dim a
Static c As Long
a = 1
Randomize
'Debug.Print Rnd()
'Debug.Print Rnd()
'Debug.Print Rnd()
'Form2.Text = "1"
'aa2
'Debug.Print c
'aa2
'Debug.Print c
'aa2
'Debug.Print c
'Load Form2

Call att.addValCon(1, 1, 1, 1)

MakeConnection


End Sub
Public Function aa()

End Function
Public Static Sub aa2()

c = c + 1
Dim bb
bb = 2 + bb
Debug.Print c

End Sub

Private Sub h_Change()
MsgBox "c"
End Sub

Private Sub h1_valChange(a As Integer)
MsgBox "h1vChange"
End Sub

Private Sub VValue1_valChange(a As Integer)

End Sub
