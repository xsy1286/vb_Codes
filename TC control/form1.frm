VERSION 5.00
Object = "{6BF52A50-394A-11D3-B153-00C04F79FAA6}#1.0#0"; "wmp.dll"
Begin VB.Form Form1 
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   4995
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   8115
   Icon            =   "form1.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   4995
   ScaleWidth      =   8115
   ShowInTaskbar   =   0   'False
   Begin VB.TextBox Text2 
      Height          =   975
      Left            =   2880
      TabIndex        =   2
      Text            =   "Text2"
      Top             =   1320
      Width           =   1455
   End
   Begin VB.TextBox Text1 
      Height          =   495
      Left            =   3120
      TabIndex        =   1
      Text            =   "Text1"
      Top             =   2880
      Width           =   735
   End
   Begin VB.Timer Timer2 
      Left            =   1440
      Top             =   2040
   End
   Begin VB.Timer Timer1 
      Left            =   840
      Top             =   1920
   End
   Begin WMPLibCtl.WindowsMediaPlayer w1 
      Height          =   3600
      Left            =   690
      TabIndex        =   3
      Top             =   330
      Width           =   3675
      URL             =   ""
      rate            =   1
      balance         =   0
      currentPosition =   0
      defaultFrame    =   ""
      playCount       =   1
      autoStart       =   -1  'True
      currentMarker   =   0
      invokeURLs      =   -1  'True
      baseURL         =   ""
      volume          =   50
      mute            =   0   'False
      uiMode          =   "full"
      stretchToFit    =   0   'False
      windowlessVideo =   0   'False
      enabled         =   -1  'True
      enableContextMenu=   -1  'True
      fullScreen      =   0   'False
      SAMIStyle       =   ""
      SAMILang        =   ""
      SAMIFilename    =   ""
      captioningID    =   ""
      enableErrorDialogs=   0   'False
      _cx             =   6482
      _cy             =   6350
   End
   Begin VB.Label Label1 
      BackColor       =   &H8000000D&
      ForeColor       =   &H80000013&
      Height          =   255
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   615
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim t, a, b, d, rsttime As Long
Dim temp(1 To 2) As String
Dim mut(1 To 3) As Integer
Const appName = "timeControler"

Private Sub Form_Load()

App.TaskVisible = False

Call init_dir("TC")
 
a = Val(init_txt("TC", "1", "55")) * 60
rsttime = Val(init_txt("TC", "2", "600"))  '2.txtΪ�뵥λ

mut(1) = Val(init_txt("TC", "mt1", "200"))
mut(2) = Val(init_txt("TC", "mt2", "200"))
mut(3) = Val(init_txt("TC", "mt3", "200"))

t = 0
Timer1.Interval = 1000

Me.Height = 210
Me.Width = 555


Call setStartUp(True, appName)
End Sub


Private Sub Timer1_Timer()
t = t + 1
'Debug.Print CStr(t - rsttime) & "  "; CStr(a)
If ((t + rsttime) Mod a) = 0 And (t + rsttime) > 0 Then  '0 Mod x =0
 ' Debug.Print "form2": End
  
  b = 0
  Timer2.Interval = 1000
  temp(1) = Str(t \ 3600) + " hour(s)" + Str((t - (t \ 3600) * 3600) \ 60) + " minutes"

End If

End Sub

Private Sub Timer2_Timer()
b = b + 1

If b = 5 Then
Open "d:\Myuse\TC\m1.txt" For Binary As #1
temp(2) = Input(LOF(1), 1)
Close #1
w1.URL = temp(2)

ElseIf b = mut(1) Then
Open "d:\Myuse\TC\m2.txt" For Binary As #1
temp(2) = Input(LOF(1), 1)
Close #1
w1.URL = temp(2)

ElseIf b = mut(1) + mut(2) Then
Open "d:\Myuse\TC\m3.txt" For Binary As #1
temp(2) = Input(LOF(1), 1)
Close #1
w1.URL = temp(2)

ElseIf b = mut(1) + mut(2) + mut(3) Then
b = 0: Timer2.Interval = 0

End If

If b = 24 Then
Form2.Label1.Caption = "Please have a rest ,the computer have been turn on for" + temp(1)
Form1.Hide
Form2.Timer1.Interval = 1000
Form2.Timer2.Interval = 25: Form2.Check1.Value = 1
Form2.Show
End If

End Sub
