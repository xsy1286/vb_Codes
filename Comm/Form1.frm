VERSION 5.00
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "comctl32.ocx"
Object = "{648A5603-2C6E-101B-82B6-000000000014}#1.1#0"; "MSCOMM32.OCX"
Begin VB.Form Form1 
   Caption         =   "Comm"
   ClientHeight    =   5625
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   8385
   Icon            =   "Form1.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   5625
   ScaleWidth      =   8385
   StartUpPosition =   3  '窗口缺省
   Begin VB.CommandButton cmdProcess 
      Caption         =   "process"
      Height          =   360
      Left            =   0
      TabIndex        =   17
      Top             =   240
      Width           =   990
   End
   Begin VB.TextBox TmpText 
      Height          =   495
      Left            =   240
      TabIndex        =   16
      Top             =   1920
      Width           =   3855
   End
   Begin VB.CommandButton Command6 
      Caption         =   "Clear"
      Height          =   1095
      Left            =   7440
      TabIndex        =   15
      Top             =   480
      Width           =   615
   End
   Begin VB.CheckBox Check1 
      Caption         =   "immediately"
      Height          =   255
      Left            =   6960
      TabIndex        =   14
      Top             =   2760
      Width           =   1335
   End
   Begin VB.CommandButton Command5 
      Caption         =   "Command5"
      Height          =   495
      Left            =   7200
      TabIndex        =   13
      Top             =   4920
      Width           =   735
   End
   Begin VB.CommandButton Command4 
      Caption         =   "Command4"
      Height          =   615
      Left            =   6960
      TabIndex        =   12
      Top             =   1680
      Visible         =   0   'False
      Width           =   1335
   End
   Begin VB.CommandButton Command3 
      BackColor       =   &H0000C000&
      Height          =   375
      Left            =   3600
      Style           =   1  'Graphical
      TabIndex        =   11
      Top             =   3600
      Width           =   495
   End
   Begin VB.ComboBox Combo2 
      Height          =   300
      Left            =   960
      TabIndex        =   10
      Top             =   5160
      Width           =   975
   End
   Begin VB.ComboBox Combo1 
      Height          =   300
      Left            =   4680
      TabIndex        =   9
      Text            =   "Combo1"
      Top             =   3600
      Width           =   2055
   End
   Begin VB.TextBox Text5 
      Height          =   375
      Left            =   960
      TabIndex        =   8
      Top             =   4680
      Width           =   975
   End
   Begin VB.TextBox Text4 
      Height          =   375
      Left            =   960
      TabIndex        =   7
      Top             =   4200
      Width           =   975
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Set  This"
      Height          =   1335
      Left            =   120
      TabIndex        =   6
      Top             =   3840
      Width           =   615
   End
   Begin VB.TextBox Text3 
      Height          =   375
      Left            =   960
      TabIndex        =   5
      Top             =   3600
      Width           =   975
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Send"
      Height          =   495
      Left            =   4800
      TabIndex        =   3
      Top             =   2640
      Width           =   1815
   End
   Begin VB.TextBox Text2 
      Height          =   495
      Left            =   960
      TabIndex        =   2
      Top             =   2640
      Width           =   2535
   End
   Begin VB.TextBox Text1 
      Height          =   1215
      Left            =   960
      MultiLine       =   -1  'True
      TabIndex        =   1
      Top             =   480
      Width           =   5775
   End
   Begin MSCommLib.MSComm MSComm1 
      Left            =   5760
      Top             =   4320
      _ExtentX        =   1005
      _ExtentY        =   1005
      _Version        =   393216
      DTREnable       =   -1  'True
   End
   Begin ComctlLib.Slider Slider1 
      Height          =   675
      Left            =   2400
      TabIndex        =   0
      Top             =   4560
      Width           =   1815
      _ExtentX        =   3201
      _ExtentY        =   1191
      _Version        =   327682
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Get"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   21.75
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   435
      Left            =   3480
      TabIndex        =   4
      Top             =   0
      Width           =   720
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim bite, wei, stp As Integer
Dim jo As String
Dim comm As String
Dim i As Integer
Dim ptr As Integer
Dim midv() As Byte
Dim dep As Integer
Dim kout As Double
Dim firstRec As Integer '防止万一第一次接收时，缓存内已有大量数据
Dim numRev As Integer


Private Sub Check1_Click()
If Check1.Value = 1 Then
 Text2.Text = ""
Else:
End If
End Sub

Private Sub cmdProcess_Click()
threadVar = 0
id = CreateThread(Null, 0, AddressOf recComm, VarPtr(0), 0, id)
Dim asd As Long
'asd = CreateThread(Null, 0, AddressOf xywhXC1, VarPtr(0), 0, asd)
End Sub

Private Sub Combo1_Click() 'combo1_click是选择了下拉项  选择的值就combo1.text
On Error GoTo erhandler
Debug.Print Combo1.Text

comm = Combo1.Text
MSComm1.CommPort = Val(Mid(comm, 4, 1))   'MSComm 串口号选择，开关串口都得延时等待
delayus 500
MSComm1.PortOpen = True ': Command3.BackColor = vbGreen
delayus 500

 If MSComm1.PortOpen = True Then Command3.BackColor = vbGreen
'If MSComm1.PortOpen = False Then
'Command3.BackColor = vbRed
'Else: Command3.BackColor = vbGreen
'End If

erhandler: ' Text1.Text = CStr(Err.Number)
 Debug.Print CStr(Err.Number)
 
           If Err.Number = 8002 Then   '打开串口失败
            Command3.BackColor = vbRed: firstRec = 1
            Exit Sub
           ElseIf Err.Number = 8005 Then  '换串口得先关串口（带延时），再换（换后还要延时）
            MSComm1.PortOpen = False: firstRec = 1
            delayus 500
            Resume
           End If
           
End Sub

Private Sub Command1_Click()
If MSComm1.PortOpen = True Then
delayus 10
If MSComm1.PortOpen = True Then
MSComm1.Output = Text2.Text
End If
End If
End Sub

Private Sub Command2_Click()
 comm = Combo1.Text
jo = Combo2.Text
bite = Val(Text3.Text)
wei = Val(Text4.Text)
stp = Val(Text5.Text)
MSComm1.Settings = CStr(bite) & ",N," & CStr(wei) & "," & CStr(stp)
MSComm1.CommPort = Val(Mid(comm, 4, 1))
End Sub

Private Sub Command3_Click()
On Error GoTo erhandler
If MSComm1.PortOpen = False Then
MSComm1.PortOpen = True: Command3.BackColor = vbGreen
Else: MSComm1.PortOpen = False
     Command3.BackColor = vbRed
End If
delayus 500

If MSComm1.PortOpen = False Then
Command3.BackColor = vbRed:: firstRec = 1
Else: Command3.BackColor = vbGreen
End If

erhandler: If Err.Number = 8002 Then
            Command3.BackColor = vbRed
            Exit Sub
           End If
End Sub

Private Sub Command4_Click()
Print CStr(MSComm1.CommPort) & "   " & CStr(MSComm1.PortOpen)

End Sub

Private Sub Command5_Click()
Dim i As Integer
For i = 0 To 10
delayus 1000000
Print "1s"
Next i
End Sub

Private Sub Command6_Click()
Text1.Text = ""
End Sub

Private Sub Form_Load()

ptr = 0
firstRec = 1

On Error GoTo erhandler
Me.Left = (Screen.Width - Me.Width) / 2
Me.Top = (Screen.Height - Me.Height) / 2
For i = 1 To 9
Combo1.AddItem "COM" & CStr(i)
Next i
Combo2.AddItem "None"
Combo2.AddItem "Odd"
Combo2.AddItem "Even"
Combo2.AddItem "Mark"
Combo2.AddItem "Space"

init_dir ("Comm")
comm = init_txt("Comm", "CH", "COM3")
bite = Val(init_txt("Comm", "bite", "4800"))
wei = Val(init_txt("Comm", "wei", "8"))
stp = Val(init_txt("Comm", "stp", "1"))
jo = init_txt("Comm", "jo", "None")


Combo2.Text = jo
Text3.Text = CStr(bite)
Text4.Text = CStr(wei)
Text5.Text = CStr(stp)

MSComm1.Settings = CStr(bite) & ",N," & CStr(wei) & "," & CStr(stp)
'MSComm1.CommPort = Val(Mid(comm, 4, 1))
MSComm1.InBufferSize = 8
MSComm1.OutBufferSize = 8

MSComm1.RThreshold = 1
MSComm1.SThreshold = 1
MSComm1.InputLen = 0
'MSComm1.InputMode = comInputModeText
MSComm1.InputMode = comInputModeBinary
'If MSComm1.PortOpen = False Then MSComm1.PortOpen = True

Combo1.Text = comm  '这里也直接去Combo1_change()
MSComm1.CommPort = Val(Mid(comm, 4, 1))
MSComm1.PortOpen = True ': Command3.BackColor = vbGreen
delayus 1000

If MSComm1.PortOpen = False Then
Command3.BackColor = vbRed
Else: Command3.BackColor = vbGreen
End If

MSComm1.InBufferCount = 0

erhandler: ' Text1.Text = CStr(Err.Number)
          If Err.Number = 8002 Then
            
            Command3.BackColor = vbRed
            Exit Sub
           End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
txtPrint "Comm", "CH", comm
txtPrint "Comm", "bite", CStr(bite)
txtPrint "Comm", "wei", CStr(wei)
txtPrint "Comm", "stp", CStr(stp)
txtPrint "Comm", "jo", jo

End Sub

Private Sub MSComm1_OnComm()




Select Case MSComm1.CommEvent
Case comEvReceive

threadVar = 1
'If firstRec = 1 Then
'MSComm1.InBufferCount = 0
'firstRec = 0
'Else
'
'End If
'numRev = numRev + 1
'If numRev = 300 Then numRev = 0: Form1.Text1.Text = ""
'Form1.Text1.Text = Form1.Text1.Text + MSComm1.Input


MSComm1.InBufferCount = 0

End Select


End Sub

Private Sub Text2_KeyUp(KeyCode As Integer, Shift As Integer)

If (Check1.Value) Then
MSComm1.Output = Mid(Text2.Text, 1, 1)
Text2.Text = ""
End If
End Sub
Function ConvertHexChr(str As String) As Integer
On Error GoTo this
Dim test As Integer
test = Asc(str)
If test >= Asc("0") And test <= Asc("9") Then
test = test - Asc("0")
ElseIf test >= Asc("a") And test <= Asc("f") Then
test = test - Asc("a") + 10
ElseIf test >= Asc("A") And test <= Asc("F") Then
test = test - Asc("A") + 10
 Else
this: test = -1       '出错信息
End If
ConvertHexChr = test
End Function

