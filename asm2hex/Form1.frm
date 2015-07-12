VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form Form1 
   Caption         =   " ：默认16进"
   ClientHeight    =   6795
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   7710
   LinkTopic       =   "Form1"
   ScaleHeight     =   6795
   ScaleWidth      =   7710
   StartUpPosition =   3  '窗口缺省
   Begin VB.CommandButton Command5 
      Caption         =   "with goto"
      Height          =   375
      Left            =   6720
      TabIndex        =   6
      Top             =   5520
      Width           =   855
   End
   Begin VB.CommandButton Command4 
      Caption         =   "Command4"
      Height          =   375
      Left            =   3240
      TabIndex        =   5
      Top             =   1320
      Visible         =   0   'False
      Width           =   1095
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Command3"
      Height          =   495
      Left            =   1200
      TabIndex        =   4
      Top             =   1320
      Visible         =   0   'False
      Width           =   1095
   End
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   1560
      Top             =   6240
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Gene"
      Height          =   375
      Left            =   5280
      TabIndex        =   3
      Top             =   5520
      Width           =   975
   End
   Begin VB.TextBox Text2 
      Height          =   375
      Left            =   120
      TabIndex        =   2
      Top             =   5520
      Width           =   3255
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Open"
      Height          =   375
      Left            =   3720
      TabIndex        =   1
      Top             =   5520
      Width           =   1095
   End
   Begin VB.Timer Timer1 
      Left            =   720
      Top             =   6240
   End
   Begin VB.TextBox Text1 
      Height          =   5415
      Left            =   0
      MultiLine       =   -1  'True
      TabIndex        =   0
      Top             =   0
      Width           =   7695
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim lne As Long
Dim k As Integer
Dim adr As String
Dim asmcode() As String
Dim isbin As String
Dim ishex As String
Dim fste As String
Dim sece As String
Dim trde As String
Dim shw As Boolean

Private Sub Command1_Click()

CommonDialog1.ShowSave
If Len(CommonDialog1.FileName) <> 0 Then
adr = CommonDialog1.FileName
Debug.Print CommonDialog1.FileName

lne = txtline(adr)
Debug.Print (CStr(lne))

Text2.Text = adr
End If


End Sub

Private Sub Command2_Click()
Dim j As Long
j = 0
Dim binform As String
Dim hexform As String

'Const num = lne / 1  'wrong programmer

ReDim asmcode(1 To lne) As String

'''''''''
Open adr For Input As #1 ' 打开文件。
    
    Do While Not EOF(1) ' 循环至文件尾。
    j = j + 1
      Line Input #1, asmcode(j)
      Debug.Print asmcode(j)
    Loop
    
Close #1 ' 关闭文件。
'''''''''
Dim result As Boolean

For j = 1 To lne

If (asmcode(j) <> "") Then
If (Mid(asmcode(j), 1, 2) <> "//") Then

result = sep(asmcode(j), fste, sece, trde)
Debug.Print (CStr(result))
Debug.Print (fste): Debug.Print (sece): Debug.Print (trde)

If (trde = "w") Then
trde = 0
ElseIf (trde = "f") Then
trde = 1
End If

hexform = "02" & vtoh(j, 4) '

hexform = hexform & "000" & picasm2(fste, sece, trde, j)

hexform = hexform & moder(hexform)

t1in (":" & hexform)

End If
End If

Next j

'end of .hex
t1in (":00000001FF")
End Sub
Private Function moder(ByVal aout As String) As String
Dim iii As Long
iii = Val("&H" & aout)
iii = 256 - (iii Mod 256)
Debug.Print "iii= " & CStr(iii)
 moder = vtoh(iii, 2)

End Function

Private Function sep(incode As String, ByRef fst As String, ByRef sec As String, ByRef trd As String) As Boolean
Dim lens As Integer
Dim space As Integer
Dim dot As Integer
lens = Len(incode)
space = 0: dot = 0

For k = 1 To lens

If (Mid(incode, k, 1) = " ") Then
space = k
ElseIf (Mid(incode, k, 1) = ",") Then
dot = k
End If

Next k

fst = "": sec = "": trd = ""
If (space = 0) Then
fst = incode
Else
  fst = Mid(incode, 1, space - 1)
    If (dot = 0) Then
    sec = Mid(incode, space + 1, lens - space)
    Else
    sec = Mid(incode, space + 1, dot - space - 1)
    trd = Mid(incode, dot + 1, lens - dot)
    End If
End If

sep = True

End Function


Private Sub Command3_Click()
t1in ("123")
End Sub
Private Function t1in(into As String)
Text1.Text = Text1.Text & into & vbCrLf
End Function


Private Sub Command4_Click()
Command4.Caption = vtoh(256, 5)
End Sub

Private Sub Command5_Click()

    If (Form1.Text2.Text <> "") Then
        Form2.Show
    Else
        MsgBox "Empty"
    End If

End Sub

Private Sub Form_Load()
mid_Form Me

Text1.Top = 0
Text1.Left = 0
Text1.Width = Me.Width
Text1.Height = Me.Height - 1000

Text2.Top = Me.Height - 970
Command1.Top = Me.Height - 970
Command2.Top = Me.Height - 970
Command5.Top = Me.Height - 970
CommonDialog1.Filter = "asm Files (*.asm)|*.asm|H Files (*.H)|*.h|C Files (*.C)|*.c|TXT Files (*.TXT)|*.txt|All Files (*.*)|*.*"

shw = False

Timer1.Interval = 50
End Sub


Private Sub Form_Resize()

Debug.Print CStr(Me.WindowState)
End Sub

Private Sub Timer1_Timer()

If (Me.WindowState = 0) Or (Me.WindowState = 2) Then

If (Me.Height < 2000) Then Me.Height = 2000
Text1.Width = Me.Width
Text1.Height = Me.Height - 1000

Text2.Top = Me.Height - 970
Command1.Top = Me.Height - 970
Command2.Top = Me.Height - 970
Command5.Top = Me.Height - 970
End If

End Sub
