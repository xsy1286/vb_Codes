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
   Begin VB.Frame Frame1 
      Caption         =   "Frame1"
      Height          =   1815
      Left            =   480
      TabIndex        =   1
      Top             =   240
      Width           =   3735
      Begin VB.TextBox Text2 
         Height          =   375
         Left            =   1080
         TabIndex        =   3
         Text            =   "Text2"
         Top             =   1080
         Width           =   2295
      End
      Begin VB.TextBox Text1 
         Height          =   375
         Left            =   1080
         TabIndex        =   2
         Text            =   "C:\Users\Administrator\Desktop\vb_tmp\miandan.txt"
         Top             =   360
         Width           =   2295
      End
      Begin VB.Label Label2 
         Caption         =   "输出文件"
         Height          =   375
         Left            =   120
         TabIndex        =   5
         Top             =   1080
         Width           =   735
      End
      Begin VB.Label Label1 
         Caption         =   "输入文件"
         Height          =   375
         Left            =   120
         TabIndex        =   4
         Top             =   360
         Width           =   855
      End
   End
   Begin VB.Timer Timer1 
      Left            =   120
      Top             =   2280
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Command1"
      Height          =   495
      Left            =   960
      TabIndex        =   0
      Top             =   2280
      Width           =   2175
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim t As Long
Dim line_n As Long
Dim ci() As String
Dim w() As String
Private Type IDM
        nam As String
        subNum As Integer
        sub1 As String
        sub2 As String
        sub3 As String
        sub4 As String
        sub5 As String
        sub6 As String
End Type
Dim tle As String
Dim p(0 To 2000) As IDM
Dim pN As Long '人数以1开始算起
Private tmpLine As Long

Private Sub Command1_Click()
Dim r As Long
Dim adr1 As String

adr1 = Text1.Text
line_n = txtline(adr1)

ReDim ci(line_n)


Call rdTxt("C:\Users\Administrator\Desktop\vb_tmp", "miandan", ci, line_n)
Debug.Print r

Timer1.Interval = 10
Me.Caption = "请稍等"
Command1.Enabled = False

Call deg
Call linedeg
End Sub

Private Sub Form_Load()
t = 0
tmpLine = 0
pN = 0
End Sub

Private Sub Timer1_Timer()
t = t + 1
Dim r As Integer
r = InStr(ci(tmpLine), "班")

If r <> 0 Then
   tle = Mid(ci(tmpLine), 1, r - 1)
   
   Dim e As String
   e = Mid(tle, Len(tle), 1)
   
   Select Case e
    Case "一":
     tle = Mid(tle, 1, Len(tle) - 1) + "①"
     Case "二":
     tle = Mid(tle, 1, Len(tle) - 1) + "②"
      Case "三":
     tle = Mid(tle, 1, Len(tle) - 1) + "③"
      Case "四":
     tle = Mid(tle, 1, Len(tle) - 1) + "④"
      Case "五":
     tle = Mid(tle, 1, Len(tle) - 1) + "⑤"
   End Select
      
   
Else

Dim MyArray
MyArray = Split(ci(tmpLine), " ", -1, 1)

Dim i As Integer
For i = LBound(MyArray) To UBound(MyArray)
    Call index(MyArray(i), tle)
Next i

End If


tmpLine = tmpLine + 1

If tmpLine = line_n Then
Call finish
Timer1.Interval = 0
End If

End Sub

Private Function index(ByVal pName As String, ByVal tle As String) As String
Dim i As Long
Dim r As String
r = ""
    For i = 1 To pN
       If p(i).nam = pName Then
            p(i).subNum = p(i).subNum + 1
            Select Case p(i).subNum
          '  Case 1:
           '     p(i).sub1 = tle
            Case 2:
                p(i).sub2 = tle
            Case 3:
                p(i).sub3 = tle
            Case 4:
                p(i).sub4 = tle
            Case 5:
                p(i).sub5 = tle
            Case 6:
                p(i).sub6 = tle
            End Select
            r = "第" & i & "个的第" & p(i).subNum & "课程"
            
            Exit Function
        End If
        
    Next i
    If r = "" Then
         pN = pN + 1
         p(pN).nam = pName
         p(pN).subNum = 1
         p(pN).sub1 = tle
         
       r = "new insert"
    End If
    
    index = r
    
End Function

Private Sub finish()
Dim i, k, j
ReDim w(pN)
    For i = 1 To pN
        w(i - 1) = p(i).nam
        
            k = p(i).subNum
            
          For j = 1 To k
          
              Select Case j
                 Case 1:
                   w(i - 1) = w(i - 1) & " " & p(i).sub1
                 Case 2:
                      w(i - 1) = w(i - 1) & " " & p(i).sub2
                 Case 3:
                       w(i - 1) = w(i - 1) & " " & p(i).sub3
                 Case 4:
                       w(i - 1) = w(i - 1) & " " & p(i).sub4
                 Case 5:
                       w(i - 1) = w(i - 1) & " " & p(i).sub5
                 Case 6:
                       w(i - 1) = w(i - 1) & " " & p(i).sub6
                 End Select
                 
           Next j
           
           
       Next i


Call wrTxt("C:\Users\Administrator\Desktop\vb_tmp", "miandan_trans", w, pN)
    Me.Caption = "已完成"
    Call Form_Load
    Command1.Enabled = True

End Sub

Private Sub deg()
Dim i As Long
Dim r As String
Debug.Print "start instr 班"
For i = 0 To (line_n - 1)

r = InStr(ci(i), "班")
If r <> 0 Then Debug.Print "i=" & i; " ,r=" & r

Next i

End Sub

Private Sub linedeg()
Dim i As Long
Dim r As String
Const lineNumber = 6

Debug.Print "start linedeg"
For i = 1 To Len(ci(lineNumber))
Debug.Print "第" & i&; "个是" & Mid(ci(lineNumber), i, 1)
Next i

End Sub
