VERSION 5.00
Begin VB.Form main 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "输入记录文件名"
   ClientHeight    =   1035
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   3735
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1035
   ScaleWidth      =   3735
   StartUpPosition =   3  '窗口缺省
   Begin VB.TextBox Text1 
      Height          =   375
      Left            =   1080
      TabIndex        =   2
      Top             =   120
      Width           =   1455
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Cancel"
      Height          =   375
      Left            =   2040
      TabIndex        =   1
      Top             =   600
      Width           =   1215
   End
   Begin VB.CommandButton Command1 
      Caption         =   "OK"
      Height          =   375
      Left            =   480
      TabIndex        =   0
      Top             =   600
      Width           =   975
   End
End
Attribute VB_Name = "main"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public appName As String
Public txtnam As String

Private Sub Command1_Click()

    Dim r As Long

    If Text1.Text = "" Then
        MsgBox "请输入文件名"
    ElseIf Dir("D:\Myuse\" & appName & "\" & Text1.Text) <> "" Then
        MsgBox "文件名重复"
    Else
        r = txtCrt(appName, Text1.Text)

        If r = 1 Then
            txtnam = Text1.Text
            Load Form1
            Me.Hide
        End If
    End If

End Sub

Private Sub Command2_Click()
    End
End Sub

Private Sub Form_Load()
    Call mid_Form(Me)
    appName = "preControlCur"
    init_dir (appName)
End Sub
