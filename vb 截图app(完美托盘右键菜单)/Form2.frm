VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form Form2 
   AutoRedraw      =   -1  'True
   BorderStyle     =   0  'None
   Caption         =   "ScreeforShot"
   ClientHeight    =   7410
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   11640
   LinkTopic       =   "Form2"
   ScaleHeight     =   7410
   ScaleWidth      =   11640
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  '����ȱʡ
   Begin VB.CommandButton Command1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "�ر�"
      Height          =   375
      Left            =   0
      MaskColor       =   &H000080FF&
      Style           =   1  'Graphical
      TabIndex        =   11
      Top             =   360
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.Timer Timer1 
      Left            =   7080
      Top             =   4200
   End
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   5760
      Top             =   3720
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.CheckBox Check1 
      Alignment       =   1  'Right Justify
      BackColor       =   &H8000000B&
      Caption         =   "������ͼƬ"
      Height          =   255
      Left            =   960
      TabIndex        =   10
      Top             =   60
      Width           =   1695
   End
   Begin VB.ComboBox Combo1 
      Height          =   300
      Left            =   4680
      TabIndex        =   8
      Text            =   ".bmp"
      Top             =   360
      Visible         =   0   'False
      Width           =   900
   End
   Begin VB.TextBox Text2 
      Height          =   375
      Left            =   960
      TabIndex        =   7
      Top             =   960
      Visible         =   0   'False
      Width           =   6000
   End
   Begin VB.CommandButton Command5 
      BackColor       =   &H00FFFF00&
      Caption         =   "������а�"
      Height          =   300
      Left            =   5760
      Style           =   1  'Graphical
      TabIndex        =   6
      Top             =   2160
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.TextBox Text1 
      Height          =   375
      Left            =   2640
      TabIndex        =   4
      Top             =   360
      Visible         =   0   'False
      Width           =   1695
   End
   Begin VB.DirListBox Dir1 
      Height          =   1770
      Left            =   960
      TabIndex        =   3
      Top             =   2280
      Visible         =   0   'False
      Width           =   4095
   End
   Begin VB.CommandButton Command4 
      BackColor       =   &H00FFFF00&
      Caption         =   "ȡ��"
      Height          =   300
      Left            =   6960
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   2880
      Visible         =   0   'False
      Width           =   500
   End
   Begin VB.CommandButton Command3 
      BackColor       =   &H00FFFF00&
      Caption         =   "ȷ��"
      Height          =   300
      Left            =   6480
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   2880
      Visible         =   0   'False
      Width           =   500
   End
   Begin VB.CommandButton Command2 
      BackColor       =   &H00FFFFFF&
      Caption         =   "ȡ����ͼ"
      Height          =   375
      Left            =   0
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   0
      Width           =   855
   End
   Begin VB.Line Line4 
      X1              =   2760
      X2              =   6000
      Y1              =   7080
      Y2              =   7080
   End
   Begin VB.Line Line3 
      X1              =   5760
      X2              =   5760
      Y1              =   4680
      Y2              =   6600
   End
   Begin VB.Line Line2 
      X1              =   2880
      X2              =   5520
      Y1              =   4560
      Y2              =   4560
   End
   Begin VB.Line Line1 
      Visible         =   0   'False
      X1              =   2520
      X2              =   2520
      Y1              =   4320
      Y2              =   6960
   End
   Begin VB.Label Label2 
      Caption         =   "Label2"
      BeginProperty Font 
         Name            =   "����"
         Size            =   14.25
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1215
      Left            =   480
      TabIndex        =   9
      Top             =   1680
      Visible         =   0   'False
      Width           =   2535
   End
   Begin VB.Label Label1 
      Caption         =   "��ַ������"
      BeginProperty Font 
         Name            =   "����"
         Size            =   14.25
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   960
      TabIndex        =   5
      Top             =   360
      Visible         =   0   'False
      Width           =   1575
   End
   Begin VB.Shape Shape1 
      BorderColor     =   &H00000000&
      BorderWidth     =   2
      Height          =   975
      Left            =   5520
      Top             =   1920
      Width           =   5895
   End
End
Attribute VB_Name = "Form2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'Integer ��������Integer �����洢Ϊ 16λ��2 ���ֽڣ�����ֵ��ʽ���䷶ΧΪ -32,768 �� 32,767 ֮�䡣Integer �����������ַ��ǰٷֱȷ��� (%)��Long ��������
'Long�������ͣ������洢Ϊ 32 λ��4 ���ֽڣ��з��ŵ���ֵ��ʽ���䷶Χ�� -2,147,483,648 �� 2,147,483,647��Long �����������ַ�Ϊ�ͺ� (&)��

Dim x1&, x2&, y1&, y2 As Long

Dim p1 As POINTAPI
Dim p2 As POINTAPI
Dim p3 As POINTAPI

Dim changedown As Integer

Private drawls As Integer
Private Type POINTAPI
    X As Long
    Y As Long
End Type
Private Declare Function GetCursorPos Lib "user32" (lpoint As POINTAPI) As Long
Private Declare Function SetWindowPos Lib "user32" (ByVal hwnd As Long, ByVal hWndInsertAfter As Long, ByVal X As Long, ByVal Y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long) As Long
Const HWND_TOPMOST = -1
Const HWND_NOTOPMOST = -2
Const SWP_NOSIZE = &H1
Const SWP_NOMOVE = &H2
Const SWP_NOACTIVATE = &H10
Const SWP_SHOWWINDOW = &H40

Private Declare Function GetSystemMetrics Lib "user32" (ByVal nIndex As Long) As Long
Private Const SM_CXSCREEN = 0
Private Const SM_CYSCREEN = 1

Private Declare Sub keybd_event Lib "user32" (ByVal bVk As Byte, ByVal bScan As Byte, ByVal dwFlags As Long, ByVal dwExtraInfo As Long) 'ģ�����Screen Print API
Private Const KEYEVENTF_KEYUP = &H2

Const cmd5w = 1200
Const cmd5h = 300



Sub SetFormTopmost(TheForm As Form)
SetWindowPos Me.hwnd, HWND_TOPMOST, 0, 0, 0, 0, SWP_NOACTIVATE Or SWP_SHOWWINDOW Or SWP_NOMOVE Or SWP_NOSIZE
End Sub



Private Sub Check1_Click()
If Check1.Value = 1 Then
Form1.Text2.Text = "1"
Else:
Form1.Text2.Text = "0"
End If
Open "d:\Myuse\shot\address.txt" For Output As #1
Print #1, (Form1.Text1.Text) & (Form1.Text2.Text) & (Form1.Text3.Text) & Mid(Str(dfaddress), 2) & Mid(Str(clip), 2) & Mid(Str(v6), 2)
Close #1
End Sub



Private Sub Combo1_Change()
Text2.Text = Dir1.Path & "\" & Text1.Text & Combo1.Text
End Sub

Private Sub Command1_Click()

'Call UnHooK

End

End Sub

Private Sub Command2_Click()

'Call UnHooK
If v = 0 Then
Form1.Show
Form1.Refresh
Form1.Check1.Refresh
Form1.Check2.Refresh
Form1.Check3.Refresh
Form1.Check4.Refresh
Form1.Command1.Refresh
Form1.Label1.Refresh
Form1.Label2.Refresh
Form1.Label3.Refresh
Form1.Text4.Refresh
Form1.Combo1.Refresh
End If
 Unload Form2
End Sub

Private Sub Command3_Click()

    Form1.Picture1.Refresh
    Form2.Text1.Text = Year(Date) & Month(Date) & Day(Date) & Hour(time) & Minute(time) & Second(time) & "��ͼ"
    Call UnHooK

    showComCmd False

    Form1.Picture1.Refresh

    If Form1.Check3.Value = 1 Then
    Dim ads As String
      ads = Form1.Text4.Text & Form2.Text1.Text & Form1.Combo1.Text
             If Dir(Form1.Text4.Text) <> "" Then
                 SavePicture Form1.Picture1.Image, (ads)
             Else: MsgBox "Wrong Path", 0, "Warning": GoTo OpenFile
             End If
       If v = 0 Then

        Form1.Show
        Form1.Refresh
        Form1.Check1.Refresh
        Form1.Check2.Refresh
        Form1.Check3.Refresh
        Form1.Check4.Refresh
        Form1.Command1.Refresh
        Form1.Label1.Refresh
        Form1.Label2.Refresh
        Form1.Label3.Refresh
        Form1.Text4.Refresh
        Form1.Combo1.Refresh
    End If

    Form2.Hide
 
        If Form2.Check1.Value = 1 Then Shell "explorer " & ads, 1  'shell��������C��  Shell "explorer ·��"
 
    Unload Form2
    
    Exit Sub
    End If

OpenFile:
    CommonDialog1.ShowSave

    If Len(Form2.CommonDialog1.FileName) <> 0 Then
        SavePicture Form1.Picture1.Image, (Form2.CommonDialog1.FileName)
    
AfterSave:
    If v = 0 Then

        Form1.Show
        Form1.Refresh
        Form1.Check1.Refresh
        Form1.Check2.Refresh
        Form1.Check3.Refresh
        Form1.Check4.Refresh
        Form1.Command1.Refresh
        Form1.Label1.Refresh
        Form1.Label2.Refresh
        Form1.Label3.Refresh
        Form1.Text4.Refresh
        Form1.Combo1.Refresh
    End If

    Form2.Hide
 
        If Form2.Check1.Value = 1 Then Shell "explorer " & Form2.CommonDialog1.FileName, 1  'shell��������C��  Shell "explorer ·��"
 
    Unload Form2
    Else:
    'Call hook
 
    End If
 
End Sub

Private Sub Command4_Click()
d = 0
shapeexist = False
Line1.Visible = False
Line2.Visible = False
Line3.Visible = False
Line4.Visible = False


showComCmd (False)

Me.MousePointer = 0

dn = 0
End Sub

Private Sub Command5_Click()
    clip = 1
    Call dl
    Unload Form2
End Sub

Private Sub Dir1_Change()
Text2.Text = Dir1.Path & "\" & Text1.Text & Combo1.Text
End Sub


Private Sub Form_Load()
Dim w, h As Integer

If m_debug = 1 Then
w = GetSystemMetrics(SM_CXSCREEN) * Screen.TwipsPerPixelX / 2 'screen.width�����ô����
h = GetSystemMetrics(SM_CYSCREEN) * Screen.TwipsPerPixelY / 2
Else
w = GetSystemMetrics(SM_CXSCREEN) * Screen.TwipsPerPixelX 'screen.width�����ô����
h = GetSystemMetrics(SM_CYSCREEN) * Screen.TwipsPerPixelY
End If

Debug.Print CStr(h)


On Error GoTo Errlog

'keybd_event vbKeySnapshot, 0&, 0&, 0&  'ģ��ScreePrint��ȫ����ͼ�����а�'�������޷����⼫Ʒ�ɳ�5���к�GetSystemMetrics(SM_CXSCREEN)��ȫ��
'DoEvents



BitBlt Form2.hdc, 0, 0, w, h, GetDC(0), 0, 0, vbSrcCopy  '������ȫ������򿪺��������

Form2.Left = 0
Form2.Top = 0
Form2.Width = w
Form2.Height = h

If m_debug = 0 Then
' ��������Ϊ������ǰ
SetFormTopmost Form2
End If

CommonDialog1.Filter = "Bmp Files (*.BMP)|*.bmp|Jpg Files (*.JPG)|*.jpg|Png Files (*.PNG)|*.png|All Files (*.*)|*.*"

Command5.Width = cmd5w
Command5.Height = cmd5h

Command3.Visible = False
Shape1.Visible = False

Command2.Visible = True
Command4.Visible = False
Text1.Visible = False
Text2.Visible = False
Dir1.Visible = False
Label1.Visible = False
Command5.Visible = False
Line1.Visible = 0
Line2.Visible = 0
Line3.Visible = 0
Line4.Visible = 0

'�û��Զ���

Combo1.Visible = False
Combo1.AddItem (".jpg")
Combo1.AddItem (".png")
Combo1.AddItem (".bmp")

'��ֹ��form2 loadǰ �ᵽform2��� ����������form2 load �� form2.show
Form2.Check1.Value = Val(Form1.Text2.Text)
dn = 0
tt = 0
shapeexist = False
drawls = 0
 
'Call hook
Me.Show
Exit Sub
Errlog:
Call whenErr(Err.number, "shot", "Clip")
     If Err.number = 521 Then Resume  '���а�ʹ�ñ���on error
End Sub
Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 27 Then Call Cancelshoot
End Sub
Private Sub Check1_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 27 Then Call Cancelshoot
End Sub
Private Sub Command3_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 27 Then Call Cancelshoot
End Sub
Private Sub Command4_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 27 Then Call Cancelshoot
End Sub

Private Sub Command2_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 27 Then Call Cancelshoot
End Sub
'''


Private Sub Form_Unload(Cancel As Integer)

If Form1.Check2.Value = 1 Then
    Form1.Hide
Else

    Form1.Show
    If v = 0 Then
        Form1.Show
        Form1.Refresh
        Form1.Check1.Refresh
        Form1.Check2.Refresh
        Form1.Check3.Refresh
        Form1.Check4.Refresh
        Form1.Command1.Refresh
        Form1.Label1.Refresh
        Form1.Label2.Refresh
        Form1.Label3.Refresh
        Form1.Text4.Refresh
        Form1.Combo1.Refresh
    End If
End If
 'Call UnHooK

End Sub

Private Sub Text1_Change()
    Text2.Text = Dir1.Path & "\" & Text1.Text & Combo1.Text
End Sub


'Private Sub Label1_DblClick()
'If dn = 1 Then Call dl
'End Sub
'Private Sub Label2_DblClick()
'If dn = 1 Then Call dl
'End Sub
Private Sub Form_DblClick() '˫�������˵����¼�
'If dn = 1 Then Call dl
End Sub
'Private Sub Text1_DblClick()
'If dn = 1 Then Call dl
'End Sub
'
'Private Sub Text2_DblClick()
'If dn = 1 Then Call dl
'End Sub

Private Sub Form_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)

    showComCmd (False)

If Me.MousePointer <> 0 Then
    changedown = 1
    Timer1.Interval = 20
    Timer1.Enabled = True
    'p1.x = x: p1.y = y
Else
    x1 = X: y1 = Y
    drawls = 1
End If

End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

If drawls = 1 Then
    x2 = X: y2 = Y
    Call LineDraw
    
ElseIf shapeexist = True And changedown = 0 Then '�˴����ж�����x1,y1Ϊ���Ͻǣ�x2,y2Ϊ���½�
    Dim dxl#, dxr#, dyt#, dye As Double
    dxl = X - x1
    dxr = X - x2
    dyt = Y - y1
    dye = Y - y2
       
       If dxl > -70 And dxl < 70 And dyt > 70 And dye < -70 Then  'left
       Me.MousePointer = 9: changeshape = 1
       ElseIf dxr > -70 And dxr < 70 And dyt > 70 And dye < -70 Then  'right
       Me.MousePointer = 9: changeshape = 2
       ElseIf dyt > -70 And dyt < 70 And dxl > 70 And dxr < -70 Then  'top
       Me.MousePointer = 7: changeshape = 3
       ElseIf dye > -70 And dye < 70 And dxl > 70 And dxr < -70 Then 'end
       Me.MousePointer = 7: changeshape = 4
    
       ElseIf Sqr(dyt * dyt + dxl * dxl) < 70 Then
       Me.MousePointer = 8: changeshape = 5
       ElseIf Sqr(dyt * dyt + dxr * dxr) < 70 Then
       Me.MousePointer = 6: changeshape = 6
       ElseIf Sqr(dye * dye + dxl * dxl) < 70 Then
       Me.MousePointer = 6: changeshape = 7
       ElseIf Sqr(dye * dye + dxr * dxr) < 70 Then
       Me.MousePointer = 8: changeshape = 8
       Else
       Me.MousePointer = 0
       End If

End If


End Sub

Private Sub Form_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)

    If drawls = 1 Then
   
        Debug.Print "drawl=1"
        x2 = X: y2 = Y
        Call lineshape(x1, y1, x2, y2)
        
        drawls = 0
        
        
        Call lineshoot

    ElseIf changedown = 1 Then
        changedown = 0
        
        Call lineshape(x1, y1, x2, y2)

        Timer1.Enabled = False
        
        Call order(x1, x2)
        Call order(y1, y2)
        
        Call lineshoot
    End If

    Dim templong As Long

End Sub

Private Sub Timer1_Timer() 'For when Correcting Shape by Mouse

    'On Error Resume Next
    If changedown = 1 Then

        Dim l As Long

        l = GetCursorPos(p2) '�˴�������ֱַ����API
        p2.X = p2.X * Screen.TwipsPerPixelX
        p2.Y = p2.Y * Screen.TwipsPerPixelY

        Select Case changeshape

            Case 1
                x1 = p2.X

            Case 2
                x2 = p2.X

            Case 3
                y1 = p2.Y

            Case 4
                y2 = p2.Y

            Case 5
                x1 = p2.X
                y1 = p2.Y

            Case 6
                y1 = p2.Y
                x2 = p2.X

            Case 7
                y2 = p2.Y
                x1 = p2.X

            Case 8
                x2 = p2.X
                y2 = p2.Y
        End Select

        Call LineDraw
    End If

End Sub



Private Sub lineshape(ByVal x1 As Long, ByVal y1 As Long, ByVal x2 As Long, ByVal y2 As Long) 'lineshape��ʾ����Call order
    '����ںڿ��ڣ��ڿ���Ϊ��ͼ���������ڿ�
'    x1 = x1 - Screen.TwipsPerPixelX: x2 = x2 + Screen.TwipsPerPixelX
'    y1 = y1 - Screen.TwipsPerPixelY: y2 = y2 + Screen.TwipsPerPixelY
    
    'Debug.Print CStr(x1)
    'Debug.Print CStr(y1)
    'Debug.Print CStr(x2)
    'Debug.Print CStr(y2)
    
    Line1.x1 = x1: Line1.y1 = y1
    Line1.x2 = x2: Line1.y2 = y1
    
    Line2.x1 = x1: Line2.y1 = y1
    Line2.x2 = x1: Line2.y2 = y2
    
    Line3.x1 = x2: Line3.y1 = y1
    Line3.x2 = x2: Line3.y2 = y2
    
    Line4.x1 = x1: Line4.y1 = y2
    Line4.x2 = x2: Line4.y2 = y2

End Sub

Private Sub LineDraw()

    Call lineshape(x1, y1, x2, y2)
    
    If shapeexist = False Then
        Line1.Visible = True
        Line2.Visible = True
        Line3.Visible = True
        Line4.Visible = True
        shapeexist = True
    End If
    
End Sub

Private Sub lineshoot() 'lineshoot���ı����x1,x2,y1,y2����
On Error GoTo Errlog

    Form2.Refresh
    Form1.Picture1.Cls 'ע��������ػ�

    showComCmd (False)

Dim x1p&, x2p&, y1p&, y2p As Long
      x1p = x1 / Screen.TwipsPerPixelX
      x2p = x2 / Screen.TwipsPerPixelX
      y1p = y1 / Screen.TwipsPerPixelY
      y2p = y2 / Screen.TwipsPerPixelY
      'module ��ע��Form1�������ҪForm dot

If m_debug = 1 Then
'Dim a As Integer
'a = x1p
'Call whenErr(a, "shot", "xget")
'a = x2p
'Call whenErr(a, "shot", "xget")
End If

Call order(x1p, x2p)
Call order(y1p, y2p)

If m_debug = 1 Then
'Call whenErr(CInt(x1p), "shot", "xget")
'Call whenErr(CInt(x2p), "shot", "xget")
End If

Const offset = 3

l = (x2p - x1p) + offset 'offset��֤�ںڿ���
h = (y2p - y1p) + offset
Form1.Picture1.Width = l
Form1.Picture1.Height = h

Form2.Label2.Caption = Str(l) + "   " + Str(h)
                            '��Ϊ���ص����
                               '��        ��
BitBlt Form1.Picture1.hdc, (0 - x1p - 1), (0 - y1p - 1), l + 1300, h + 1300, GetDC(0), 0, 0, vbSrcCopy
'Debug.Print "r=" & Str(0 - x1) & Str(0 - y1)
'Form1.Picture1.PaintPicture Form2.Picture, 0, 0, Form1.Picture1.Width, Form1.Picture1.Height, x1  , y1  , Screen.Width, Screen.Height, vbSrcCopy



Call UIcmd(x2p * Screen.TwipsPerPixelX, y2p * Screen.TwipsPerPixelY)
  
dn = 1


showComCmd (True)


Errlog:
    Call whenErr(Err.number, "shot", "lineshoot()")
    
End Sub
Private Sub UIcmd(ByVal X As Long, ByVal Y As Long)
Const gap = 20
  Form2.Command3.Left = X - 1000: Form2.Command3.Top = Y - 300
  Form2.Command4.Left = X - 500: Form2.Command4.Top = Y - 300
  Form2.Command5.Left = X - 1000 - cmd5w - gap: Form2.Command5.Top = Y - 300
End Sub

Private Sub showComCmd(b As Boolean)

    Form2.Command3.Visible = b
    Form2.Command4.Visible = b
    Form2.Command5.Visible = b
End Sub
