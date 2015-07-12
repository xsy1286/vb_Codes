VERSION 5.00
Begin VB.Form Aero 
   BorderStyle     =   0  'None
   Caption         =   "Skin Vista Aero The mere idea"
   ClientHeight    =   8100
   ClientLeft      =   60
   ClientTop       =   0
   ClientWidth     =   10830
   Icon            =   "Aero.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   8100
   ScaleWidth      =   10830
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  '窗口缺省
   WindowState     =   2  'Maximized
   Begin Vista_Aero.aicAlphaImage Command1 
      Height          =   315
      Left            =   840
      ToolTipText     =   "退出"
      Top             =   7200
      Width           =   315
      _ExtentX        =   556
      _ExtentY        =   556
      Image           =   "Aero.frx":AE5A
      Props           =   5
   End
   Begin Vista_Aero.aicAlphaImage AlphaImage1 
      Height          =   3600
      Left            =   2895
      Top             =   1275
      Width           =   5400
      _ExtentX        =   9525
      _ExtentY        =   6350
      Image           =   "Aero.frx":B2F7
   End
   Begin Vista_Aero.aicAlphaImage ucAlphaImage4 
      Height          =   540
      Left            =   8820
      ToolTipText     =   "最小化"
      Top             =   390
      Width           =   690
      _ExtentX        =   1217
      _ExtentY        =   953
      Image           =   "Aero.frx":1EF51
      Props           =   5
   End
   Begin Vista_Aero.aicAlphaImage ucAlphaImage2 
      Height          =   540
      Left            =   9600
      ToolTipText     =   "关闭"
      Top             =   390
      Width           =   945
      _ExtentX        =   1667
      _ExtentY        =   953
      Image           =   "Aero.frx":1FAD9
      Props           =   5
   End
   Begin Vista_Aero.aicAlphaImage ucAlphaImage5 
      Height          =   450
      Left            =   9210
      ToolTipText     =   "最大化"
      Top             =   420
      Width           =   690
      _ExtentX        =   1217
      _ExtentY        =   794
      Image           =   "Aero.frx":208F6
      Props           =   5
   End
   Begin VB.Image Image1 
      Height          =   285
      Left            =   9765
      Top             =   520
      Width           =   615
   End
   Begin VB.Image Image2 
      Height          =   285
      Left            =   9000
      Top             =   520
      Width           =   435
   End
   Begin VB.Image Image3 
      Height          =   285
      Left            =   9375
      Top             =   520
      Width           =   435
   End
   Begin Vista_Aero.aicAlphaImage Picture1 
      Height          =   8250
      Left            =   60
      Top             =   0
      Width           =   11250
      _ExtentX        =   19844
      _ExtentY        =   14552
      Image           =   "Aero.frx":290C7
   End
End
Attribute VB_Name = "Aero"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'******************************
'  源码学习下载www.lvcode.com
'    欢迎分享源码给Love代码
'******************************
'************************
'* Karmba_a@hotmail.com *
'*      MSLE 2008       *
'************************
Private X1 As Integer, Y1 As Integer

Private Sub AlphaImage1_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
On Error Resume Next

End Sub
'鼠标进入图片淡入淡出
Private Sub AlphaImage1_MouseEnter()
    AlphaImage1.ShadowEnabled = False
    AlphaImage1.grayScale = aiNoGrayScale
  '  AlphaImage2.FadeInOut 100
End Sub

'鼠标退出图片淡入淡出
Private Sub AlphaImage1_MouseExit()
    AlphaImage1.ShadowEnabled = True
'    AlphaImage2.FadeInOut 0
End Sub

Private Sub Command1_Click(ByVal Button As Integer)
End
End Sub

Private Sub Form_Click()
On Error Resume Next
        Call SystrayOn(Me, "这是最新 Vista Aero 皮肤界面示例")
        Call SetForegroundWindow(Me.hWnd)
        Me.Hide
        Me.WindowState = 1
End Sub

Private Sub Form_Load()
On Error Resume Next
AlphaImage2.Opacity = 0&
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
On Error Resume Next
    Static lngMsg As Long
    Dim blnflag As Boolean, lngResult As Long
    
    lngMsg = X / Screen.TwipsPerPixelX
    If blnflag = False Then
        blnflag = True
        Select Case lngMsg
        Case WM_LBUTTONCLK      '单击左键弹出
            Call SystrayOff(Me)
            Call SetForegroundWindow(Me.hWnd)
            Me.WindowState = 2
            TakeScreenShot Me, ""
            Me.Show
        Case WM_LBUTTONDBLCLK   '双击左键显示窗体
            PopupMenu mnuRestore
        End Select
    End If
    ucAlphaImage2.Visible = False
    ucAlphaImage4.Visible = False
    ucAlphaImage5.Visible = False
End Sub

Private Sub Form_Resize()
    TakeScreenShot Me, ""
    Command1.Visible = True
    Picture1.Visible = True
End Sub

Private Sub Image1_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    ucAlphaImage5.Visible = False
    ucAlphaImage4.Visible = False
    ucAlphaImage2.ZOrder 0
    ucAlphaImage2.Visible = True
End Sub

Private Sub Image2_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    ucAlphaImage5.Visible = False
    ucAlphaImage2.Visible = False
    ucAlphaImage4.ZOrder 0
    ucAlphaImage4.Visible = True
End Sub
Private Sub Image3_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    ucAlphaImage2.Visible = False
    ucAlphaImage4.Visible = False
    ucAlphaImage5.ZOrder 0
    ucAlphaImage5.Visible = True
End Sub

Private Sub Picture1_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Picture1.ZOrder 0
    X1 = X
    Y1 = Y
   ' Label2.ZOrder 0
    Command1.ZOrder 0
    AlphaImage1.ZOrder 0
   ' AlphaImage2.ZOrder 0
End Sub

Private Sub Picture1_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = 1 Then
        Picture1.Left = Picture1.Left + X - X1
        Picture1.Top = Picture1.Top + Y - Y1
        Command1.Left = Command1.Left + X - X1
        Command1.Top = Command1.Top + Y - Y1
        AlphaImage1.Left = AlphaImage1.Left + X - X1
        AlphaImage1.Top = AlphaImage1.Top + Y - Y1
      '  AlphaImage2.Left = AlphaImage2.Left + X - X1
   '     AlphaImage2.Top = AlphaImage2.Top + Y - Y1
        Image1.Left = Image1.Left + X - X1
        Image1.Top = Image1.Top + Y - Y1
        Image2.Left = Image2.Left + X - X1
        Image2.Top = Image2.Top + Y - Y1
        Image3.Left = Image3.Left + X - X1
        Image3.Top = Image3.Top + Y - Y1
        ucAlphaImage2.Left = ucAlphaImage2.Left + X - X1
        ucAlphaImage2.Top = ucAlphaImage2.Top + Y - Y1
        ucAlphaImage4.Left = ucAlphaImage4.Left + X - X1
        ucAlphaImage4.Top = ucAlphaImage4.Top + Y - Y1
        ucAlphaImage5.Left = ucAlphaImage5.Left + X - X1
        ucAlphaImage5.Top = ucAlphaImage5.Top + Y - Y1
     '   Label2.Left = Label2.Left + X - X1
      '  Label2.Top = Label2.Top + Y - Y1
    End If
        ucAlphaImage2.Visible = False
        ucAlphaImage4.Visible = False
        ucAlphaImage5.Visible = False
     '   Label2.ZOrder 0
        Command1.ZOrder 0
        AlphaImage1.ZOrder 0
     '   AlphaImage2.ZOrder 0
End Sub

Private Sub Picture1_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Dim a As Integer, B As Integer, A1 As Integer, b1 As Integer
    a = Val(Left(Picture1.Tag, 7))
    B = Val(Right(Picture1.Tag, 7))
    A1 = Picture1.Left
    b1 = Picture1.Top
    If a + 300 > A1 And a - 300 < A1 And B + 300 > b1 And B - 300 < b1 Then
        Picture1.Left = a
        Picture1.Top = B
    End If
    B = 1
    For a = 0 To Sqa * Sqb - 1
        A1 = Val(Left(Picture1.Tag, 7))
        b1 = Val(Right(Picture1.Tag, 7))
        If A1 <> Picture1.Left Or b1 <> Picture1.Top Then
            B = 0
        End If
    Next a
        Image1.ZOrder 0
        Image2.ZOrder 0
        Image3.ZOrder 0
        
        
End Sub


Private Sub ucAlphaImage2_Click(ByVal Button As Integer)
End
End Sub

Private Sub ucAlphaImage4_Click(ByVal Button As Integer)
On Error Resume Next
        Call SystrayOn(Me, "这是最新 Vista Aero 皮肤界面示例")
        Call SetForegroundWindow(Me.hWnd)
        Me.Hide
        Me.WindowState = 1
End Sub



