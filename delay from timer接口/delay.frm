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
   StartUpPosition =   3  '����ȱʡ
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
'�ӿ���
'��ʱdelay(��λ100���룩
'case n ��n�������¼�
'r=timerdelay(0) Ϊ��ں��� ��ʼִ��


'����Ϊ���뵽ͨ����
'***************************************************
Option Explicit
Dim WithEvents Timerdelay As Timer '�ؼ�����
Attribute Timerdelay.VB_VarHelpID = -1
Dim t As Long
Dim n As Integer
'***************************************************



'***************************************************
'һ�¼��뵽һ�������(�ں��ӿ���)
'***************************************************
Function delay(a As Integer)  '��λʮ��֮һ��
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
Set Timerdelay = Controls.Add("VB.Timer", "Timer��Name", Form1)
Debug.Print "1"
Timerdelay.Enabled = True

'***************************************************

'�ӿ���
'��ʱdelay(��λ100���룩
'case n ��n�������¼�

delay (10) 'case 1 ʱ��֮����
Case 1
'

delay (15)
Case 2
'

delay (50)
Case 3
'








End Select ' //���̽���

'***************************************************

'VB ��� select ���� C++ ��� switch �����ͬ
'���� ÿһ��"Case"����ʱ���� <break> �мǣ���

End Function

'***************************************************
'һ�������
'***************************************************




Private Sub Form_Load()


'�ӿ����
'***************************************************
Dim r As Long
r = timedelay(0)  'r=timerdelay(0) Ϊ��ں��� ��ʼִ��
'***************************************************


End Sub


