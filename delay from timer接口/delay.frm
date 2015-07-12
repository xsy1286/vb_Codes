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
   StartUpPosition =   3  '窗口缺省
   Begin VB.CommandButton Comma 
      Caption         =   "Command2"
      Height          =   735
      Left            =   3120
      TabIndex        =   2
      Top             =   1440
      Width           =   975
   End
   Begin VB.TextBox Text1 
      Height          =   735
      Left            =   1440
      TabIndex        =   1
      Top             =   480
      Width           =   1455
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Command1"
      Height          =   495
      Left            =   1560
      TabIndex        =   0
      Top             =   1680
      Width           =   1335
   End
   Begin VB.Timer Timer1 
      Left            =   1920
      Top             =   2280
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'接口区
'延时delay(单位100毫秒）
'case n 第n次做的事件
'r=timerdelay(0) 为入口函数 开始执行


'以下为加入到通用区
'***************************************************
Option Explicit
Dim WithEvents Timerdelay As Timer '控件名称
Attribute Timerdelay.VB_VarHelpID = -1
Dim t As Long
Dim n As Integer
'***************************************************



'***************************************************
'一下加入到一般代码区(内含接口区)
'***************************************************
Function delay(a As Integer)  '单位十分之一秒
Timerdelay.Interval = a * 100
End Function
Private Sub Timerdelay_Timer()
t = t + 1
Dim r As Long
r = timedelay(t)
Debug.Print "11"
End Sub
Function timedelay(n As Long) As Long

Select Case (n)
Case 0
Set Timerdelay = Controls.Add("VB.Timer", "Timer的Name", Form1)
Debug.Print "1"
Timerdelay.Enabled = True

'***************************************************

'接口区
'延时delay(单位100毫秒）
'case n 第n次做的事件

delay (10) 'case 1 时间之后做
Case 1
'

delay (15)
Case 2
'

delay (50)
Case 3
'








End Select ' //过程结束

'***************************************************

'VB 里的 select 语句和 C++ 里的 switch 语句相同
'不过 每一个"Case"结束时无需 <break> 切记！！

End Function

'***************************************************
'一般代码区
'***************************************************




Private Sub Form_Load()


'接口入口
'***************************************************
Dim r As Long
r = timedelay(0)  'r=timerdelay(0) 为入口函数 开始执行
'***************************************************


End Sub


