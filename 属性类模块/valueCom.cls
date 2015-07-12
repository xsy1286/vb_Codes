VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "valueCom"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Const maxControl = 15
Const maxChk = 20
Private WithEvents aForm As Form
Attribute aForm.VB_VarHelpID = -1
     Const aForm_W = 5000
Private WithEvents p As SingleValue
Attribute p.VB_VarHelpID = -1
Private WithEvents cmd1 As CommandButton, WithEvents cmd2 As CommandButton
Attribute cmd1.VB_VarHelpID = -1
Attribute cmd2.VB_VarHelpID = -1
Private WithEvents txt1 As TextBox
Attribute txt1.VB_VarHelpID = -1
Private WithEvents t1 As Timer
Attribute t1.VB_VarHelpID = -1

Private v(1 To maxControl)  As SingleValue
Private c(1 To maxChk)  As CheckBox
Private chkV(1 To maxChk)  As String

Private theFormCaption As String
Private cV_title As String
Private num
Private numChk
Private singleH As Integer
Private ckH As Integer
Public Property Let caption(ByVal a As String)
    If a <> "" Then
      theFormCaption = a
      aForm.caption = a
    Else
     MsgBox "No Input"
    End If
End Property
Public Property Get caption() As String
  caption = theFormCaption
End Property
Public Property Get Number() As Integer
  Number = num
End Property

Public Property Let title(ByVal a As String)
    If a <> "" Then
      cV_title = a
    Else
     MsgBox "No Input"
    End If
End Property
Public Property Get title() As String
     title = cV_title
End Property

Public Function addChk(ByVal lb As String) As Long
    'On Error GoTo Errlog
    aForm.Width = aForm_W

    If cV_title = "" Then MsgBox "No title": Exit Function
    If numChk = maxControl Then MsgBox "Full": Exit Function
    numChk = numChk + 1
 
    Set c(numChk) = aForm.Controls.Add("VB.CheckBox", "c" & CStr(numChk))

    With c(numChk)
        .Height = 300
        .Width = 1500
   
    .Value = Val(chkV(numChk))
        Const offsetCkl = 150

        .Left = aForm_W - .Width - offsetCkl
        .Top = ckH
        ckH = ckH + 300
        
        .caption = lb
        .Visible = True
    End With

 If singleH < (ckH + cmd1.Height + cmd2.Height) Then

    Const offsetFrmH = 700
    Const div = 10

    aForm.Height = ckH + offsetFrmH + cmd1.Height + cmd2.Height
 
    cmd1.Top = ckH + offsetFrmH / div
    cmd2.Top = ckH + cmd1.Height + offsetFrmH / div
End If



    Exit Function

Errlog:

    Call whenErr(Err.Number, cV_title, "AddChk")
 
End Function

Public Function addValCon(ByVal min As Integer, _
                          ByVal max As Integer, _
                          ByVal def As Integer, _
                          ByVal lb As String) As Integer
    aForm.Width = aForm_W
    'On Error GoTo Errlog
    If cV_title = "" Then MsgBox "No title": Exit Function
    If num = maxControl Then MsgBox "Full": Exit Function
    num = num + 1

    Dim str(0 To 3) As String

    'ReDim str(4) As String
    str(0) = CStr(min)
    str(1) = CStr(max)
    str(2) = CStr(def)
    str(3) = lb
 
    Call init_txtEx(cV_title, "att" & CStr(num), str, 4)
    addValCon = Val(str(2))

    Set v(num) = aForm.Controls.Add("Object.SingleValue", "v" & CStr(num))

    With v(num)
        v(num).mTop = singleH

        If Val(str(1)) >= v(num).max Then
            v(num).max = Val(str(1))
        End If

        If Val(str(0)) <= v(num).min Then
            v(num).min = Val(str(0))
        End If

        v(num).a_val = (min + max) / 2

        v(num).min = Val(str(0))
        v(num).max = Val(str(1))
        v(num).a_val = Val(str(2))
        v(num).m_caption = str(3)
        v(num).lbl_BackStyle = 1
        v(num).stylef = h
        .mVisible = True
 
    End With
 
    singleH = singleH + v(num).mHeight
 
 If singleH >= (ckH + cmd1.Height + cmd2.Height) Then

    Const offset = 700

    aForm.Height = singleH + offset
 
    cmd1.Top = singleH - cmd1.Height '+ offset
    cmd2.Top = singleH - cmd1.Height - cmd2.Height ' + offset
End If

    Exit Function

Errlog:

    Call whenErr(Err.Number, cV_title, "AddContorlor")
 
End Function

Private Sub aForm_Load()
' Static cm1 As CommandButton
' Set cmd1 = New CommandButton
'  With cmd1
' .Visible = True
' .Left = 0
' .top = 0
' .Width = 700
' .Height = 700
' End With
' aForm.Show
Debug.Print "aFormload"

Set t1 = aForm.Controls.Add("VB.Timer", "t1")

Set cmd1 = aForm.Controls.Add("VB.CommandButton", "cmd1")
Set cmd2 = aForm.Controls.Add("VB.CommandButton", "cmd2")
Const w = 1500
Const h = 350

With cmd1
 .caption = "确定"
 .Width = w
 
 .Height = h
End With
With cmd2
 .caption = "取消"
 .Width = w
 
 .Height = h
End With
Dim i
For i = 1 To maxChk
    chkV(i) = "0"
Next

Call init_txtEx(cV_title, "Values", chkV, maxChk)

  Const leftset = 300
 cmd1.Left = aForm_W - cmd1.Width - leftset
 cmd2.Left = aForm_W - cmd2.Width - leftset
 cmd1.Visible = True
 cmd2.Visible = True
 
'aForm.Left = (GetSystemMetrics(SM_CXSCREEN) * Screen.TwipsPerPixelX - aForm.Width) / 2
'aForm.Top = (GetSystemMetrics(SM_CYSCREEN) * Screen.TwipsPerPixelY - aForm.Height) / 2
t1.Interval = 1 '刚开始的From显示只在0，0，so use timer


End Sub
Public Sub formLoad()

 With aForm
 .Visible = True
 .Left = 0
 .Top = 0
 .Width = 7000
 .Height = 7000
 End With
End Sub

Private Sub Class_Initialize()
 num = 0
 'MsgBox "1"

Set aForm = Form2
singleH = 0
ckH = 0
numChk = 0
num = 0

End Sub


Private Sub cmd2_Click()
  Unload aForm
End Sub

Private Sub cmdObject_Click()
'MsgBox p.Visible
  If p.mVisible = True Then
    p.mVisible = False
  Else
    p.mVisible = True
  End If
End Sub

Private Sub cmd1_Click()

    Dim str(0 To 3) As String

    Dim con()       As String

    ReDim con(num + 1) As String

    Dim k

   Do While num <> 0
        str(0) = v(num).min
        str(1) = v(num).max
        str(2) = CStr(v(num).a_val)
        con(num) = str(2)
        str(3) = v(num).m_caption
        wr_txtEx cV_title, "att" & CStr(num), str, 4
        num = num - 1
    Loop

    Do While numChk <> 0
    
    chkV(numChk) = CStr(c(numChk).Value)
        numChk = numChk - 1
    
    Loop
        


    wr_txtEx cV_title, "Values", chkV, maxChk
    
    Unload aForm

    'TheForm.setControl(con)
    'TheForm.setControl2(con)
End Sub

Private Sub t1_Timer()
aForm.Left = (GetSystemMetrics(SM_CXSCREEN) * Screen.TwipsPerPixelX - aForm.Width) / 2
aForm.Top = (GetSystemMetrics(SM_CYSCREEN) * Screen.TwipsPerPixelY - aForm.Height) / 2
t1.Interval = 0
End Sub

Private Sub txt1_KeyUp(KeyCode As Integer, Shift As Integer)
    Debug.Print "keyup"

    With p

        '.BackStyle = 1
        Select Case KeyCode

            Case vbKey1
                .mHeight = 1000

            Case vbKey2
                .mWidth = 1000

            Case vbKey3
                .mLeft = 3000

            Case vbKey4
                .mTop = 1500
        End Select

    End With

End Sub
