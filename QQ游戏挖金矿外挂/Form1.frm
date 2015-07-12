VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   3675
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   4635
   LinkTopic       =   "Form1"
   ScaleHeight     =   3675
   ScaleWidth      =   4635
   StartUpPosition =   3  '窗口缺省
   Begin VB.Timer Timer1 
      Left            =   1920
      Top             =   2880
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Command1"
      Height          =   1455
      Left            =   840
      TabIndex        =   0
      Top             =   1200
      Width           =   2775
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim mem() As Byte
Dim mem2() As Byte
Dim hwd As Long, casehwd() As Long, phandle As Long
Dim rect0 As RECT
Dim pBase As Long
Dim num As Integer
Const memlen = &H200000
Const mstart = &H6500000
'Const memlen = &H300000
'Const mstart = &H129F78
Const searchname = "挖金子"


Private Sub Command1_Click()
    ReDim mem(memlen)
    ReDim mem2(memlen)
    
    Dim i As Long



    ReDim casehwd(10)
    num = nameToHwndEx(searchname, casehwd, 100)
    hwd = 0
    
    For i = 0 To (num - 1)
        rect0 = WndWid(casehwd(i))
        'Call WndWid(casehwd(i), rect0)
       If rect0.Top > 0 And (rect0.Right - rect0.Left) = 800 And (rect0.Bottom - rect0.Top) = 600 Then
        hwd = casehwd(i)
        Exit For
       End If
    Next i
    If hwd = 0 Then MsgBox "未打开": Exit Sub



    Call myPrsOpenByhWnd(hwd, phandle): Debug.Print ("phandel is:" & Hex(phandle))
    
'    Call myPrsOpenByhWnd(casehwd(0), phandle): Debug.Print ("phandel is:" & Hex(phandle))
    
    
    mem() = getMem(mstart, memlen, phandle)
    Call forhandleclose(phandle)
    
    Timer1.Interval = 20
    
    Debug.Print "mouse move"
    Call SetMousePos(rect0.Left + 245, rect0.Top + 45)
    mouse_event MOUSEEVENTF_LEFTDOWN, 0, 0, 0, 0
    
    
    Call myPrsOpenByhWnd(hwd, phandle): Debug.Print ("phandel is:" & Hex(phandle))
    mem2() = getMem(mstart, memlen, phandle)
    Call forhandleclose(phandle)
    
    For i = 0 To (memlen - 1)
     If Val(mem(i)) <> 0 Then MsgBox CStr(i) & "is not zero": Exit Sub
    Next i

'    For i = 0 To (&H40000 - 1)
'       If mem(4 * i) = 0 And mem(4 * i) = mem2(4 * i) Then
'        If mem(4 * i + 1) = 0 And mem(4 * i + 1) = mem2(4 * i + 1) Then
'         If mem(4 * i + 3) < 10 And mem(4 * i + 3) = mem2(4 * i + 3) Then
'           If mem(4 * i + 2) = 0 And mem2(4 * i + 2) = 2 Then
'                pBase = 4 * i
'                Exit For
'            End If
'            End If
'            End If
'            End If
'
'    Next
'
     MsgBox "address:" & Hex(pBase)
    'MsgBox "address:" & Hex(&H6500000 + pBase)
    
'    Call SetMousePos(rect0.Left + 245, rect0.Top + 45)
'    mouse_event MOUSEEVENTF_LEFTUP, 0, 0, 0, 0
End Sub
Private Sub Form_Load()
 'Dim hToken As Long
   Call mid_Form(Me)
  Call PromotePrivileges
End Sub

Private Sub Form_Unload(Cancel As Integer)
   Call top_hWnd(hwd, False)
End Sub

Private Sub Timer1_Timer()
    
    top_hWnd (hwd)
    
End Sub
