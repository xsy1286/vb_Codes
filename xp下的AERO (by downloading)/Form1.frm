VERSION 5.00
Begin VB.Form Form1 
   BackColor       =   &H00C0C000&
   Caption         =   "AlphaImage Control by LaVolpe"
   ClientHeight    =   6045
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   6285
   LinkTopic       =   "Form1"
   ScaleHeight     =   403
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   419
   StartUpPosition =   2  'CenterScreen
   Begin VB.ComboBox cboPlayTime 
      Height          =   315
      ItemData        =   "Form1.frx":0000
      Left            =   2850
      List            =   "Form1.frx":0019
      Style           =   2  'Dropdown List
      TabIndex        =   4
      Top             =   5340
      Width           =   2610
   End
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   1000
      Left            =   60
      Top             =   45
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Caption         =   "Above icon acts like a button. Click on it during runtime"
      Height          =   465
      Index           =   4
      Left            =   3345
      TabIndex        =   6
      Top             =   3075
      Visible         =   0   'False
      Width           =   2460
   End
   Begin Project1.aicAlphaImage aicBtnImitation 
      Height          =   660
      Left            =   4830
      ToolTipText     =   "Imitates button action"
      Top             =   2310
      Width           =   690
      _ExtentX        =   1217
      _ExtentY        =   1164
      Image           =   "Form1.frx":009B
      Scaler          =   4
      Props           =   129
      ShadowOpacity   =   44
      Intensity       =   -5
      ScaleCx         =   32
      ScaleCy         =   32
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Drag File to Above or Right Click to Add Res Image"
      Height          =   405
      Index           =   5
      Left            =   255
      TabIndex        =   5
      Top             =   5280
      Width           =   2115
   End
   Begin Project1.aicAlphaImage ucPlayTime 
      Height          =   2130
      Left            =   3135
      ToolTipText     =   "Example of modifying some properties"
      Top             =   3165
      Width           =   2130
      _ExtentX        =   3757
      _ExtentY        =   3757
      Image           =   "Form1.frx":03B9
      Scaler          =   4
      Enabled         =   0   'False
      Props           =   273
      ScaleCx         =   142
      ScaleCy         =   142
   End
   Begin VB.Shape Shape1 
      Height          =   1965
      Left            =   255
      Top             =   3300
      Width           =   2085
   End
   Begin Project1.aicAlphaImage ucDragDrop 
      Height          =   1155
      Left            =   765
      ToolTipText     =   "Drag & Drop Example"
      Top             =   3705
      Width           =   1050
      _ExtentX        =   1852
      _ExtentY        =   2037
      Image           =   "Form1.frx":10356
      Scaler          =   2
      OLEdrop         =   1
   End
   Begin Project1.aicAlphaImage aicSecond 
      Height          =   1800
      Left            =   307
      Top             =   307
      Width           =   1800
      _ExtentX        =   3175
      _ExtentY        =   3175
      Image           =   "Form1.frx":1036E
      Scaler          =   4
      Angle           =   310
      Enabled         =   0   'False
      Props           =   257
      ShadowDepth     =   1
      ScaleCx         =   120
      ScaleCy         =   120
   End
   Begin Project1.aicAlphaImage aicHour 
      Height          =   1800
      Left            =   307
      Top             =   307
      Width           =   1800
      _ExtentX        =   3175
      _ExtentY        =   3175
      Image           =   "Form1.frx":10788
      Scaler          =   4
      Angle           =   220
      Enabled         =   0   'False
      Props           =   265
      ScaleCx         =   120
      ScaleCy         =   120
   End
   Begin Project1.aicAlphaImage aicMinute 
      Height          =   1800
      Left            =   307
      Top             =   307
      Width           =   1800
      _ExtentX        =   3175
      _ExtentY        =   3175
      Image           =   "Form1.frx":115D2
      Scaler          =   4
      Angle           =   90
      Enabled         =   0   'False
      Props           =   265
      ShadowColor     =   16777152
      ScaleCx         =   120
      ScaleCy         =   120
   End
   Begin Project1.aicAlphaImage aicClockDecor 
      Height          =   1245
      Left            =   660
      Top             =   660
      Width           =   1320
      _ExtentX        =   2328
      _ExtentY        =   2196
      Image           =   "Form1.frx":12475
      Enabled         =   0   'False
      Props           =   9
   End
   Begin Project1.aicAlphaImage aicClockFace 
      Height          =   1995
      Left            =   210
      ToolTipText     =   "Example of overlayed, rotated images"
      Top             =   210
      Width           =   1995
      _ExtentX        =   3519
      _ExtentY        =   3519
      Image           =   "Form1.frx":164FE
      Opacity         =   90
      HitTest         =   3
      Props           =   9
   End
   Begin Project1.aicAlphaImage aicBubble 
      Height          =   1770
      Left            =   2220
      Top             =   315
      Width           =   3165
      _ExtentX        =   5583
      _ExtentY        =   3122
      Image           =   "Form1.frx":1A98A
      Scaler          =   1
      Mirror          =   1
      Enabled         =   0   'False
      Props           =   3
      ShadowColor     =   65535
      ShadowX         =   8
      ShadowY         =   6
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Caption         =   "Bubble uses the new FadeInOut routine"
      Height          =   285
      Index           =   3
      Left            =   2265
      TabIndex        =   3
      Top             =   2010
      Visible         =   0   'False
      Width           =   3135
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Caption         =   "The bubble only appears when the mouse is over the cheetah. Cheetah changes color too"
      Height          =   435
      Index           =   2
      Left            =   2160
      TabIndex        =   2
      Top             =   0
      Visible         =   0   'False
      Width           =   3630
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Caption         =   "Above are 2 overlapped images, Both are mirrored and one is stretched"
      Height          =   480
      Index           =   1
      Left            =   2175
      TabIndex        =   1
      Top             =   2265
      Visible         =   0   'False
      Width           =   2790
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Caption         =   "Above are 5 overlapped images"
      Height          =   540
      Index           =   0
      Left            =   285
      TabIndex        =   0
      Top             =   2250
      Visible         =   0   'False
      Width           =   1650
   End
   Begin Project1.aicAlphaImage aicCheetah 
      Height          =   900
      Left            =   2880
      ToolTipText     =   "Example of Mouse Over and Fade effects"
      Top             =   735
      Width           =   2100
      _ExtentX        =   3704
      _ExtentY        =   1588
      Image           =   "Form1.frx":20A45
      Mirror          =   1
      GrayScale       =   6
      HitTest         =   3
      Props           =   129
      ShadowColor     =   4210688
      ShadowDepth     =   5
   End
   Begin VB.Menu mnuPopup 
      Caption         =   "Popup"
      Visible         =   0   'False
      Begin VB.Menu mnuSample 
         Caption         =   "Show Resource Image"
         Index           =   0
      End
      Begin VB.Menu mnuSample 
         Caption         =   "-"
         Index           =   1
      End
      Begin VB.Menu mnuSample 
         Caption         =   "Clear Image"
         Index           =   2
      End
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

' Just a sample form.  Add your own form to the project and play

Option Explicit

Private Declare Sub CopyMemory Lib "kernel32.dll" Alias "RtlMoveMemory" (ByRef Destination As Any, ByRef Source As Any, ByVal Length As Long)

Private Sub aicBtnImitation_Click(ByVal Button As Integer)
    
    ' Note that windowless usercontrols get left, right, & middle button clicks and dblClicks
    ' So, notice the Button parameter? Nice touch since VB doesn't supply it in the
    ' usercontrol's Click event
    
    ' For a button effect, we want the following properties for our control
    ' 1. Size to desired scale and then set ScaleMethod=aiLockScale
    ' 2. Increase the size enough to allow the image to shift within the control
    ' 3. Add a shadow if desired and change opacity if desired
    ' 4. Now monitor your click, mousedown, and mouseup events
    ' 5. If you are using mouse over events, then monitor the mouseExit and mouseEnter too
    If Button = vbLeftButton Then
        If Not aicBtnImitation.Tag = "MsgShown" Then
            MsgBox "Drag this message box around the form." & vbCrLf & _
                "Notice the image(s) disappearing?" & vbCrLf & _
                "This won't happen when compiled", vbInformation + vbOKOnly, "Clicked"
            aicBtnImitation.Tag = "MsgShown"
        End If
    End If
End Sub

Private Sub aicBtnImitation_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
    If (Button And vbLeftButton) = vbLeftButton Then
        ' why test for Button this way?
        ' One can hold the right button down and then also hold the left button down and
        ' the Button parameter will be (vbLeftButton Or vbRightButton)
        aicBtnImitation.OffsetImage True, 1, 1 ' shift the image
    End If
End Sub

Private Sub aicBtnImitation_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
    If (Button And vbLeftButton) = vbLeftButton Then
        aicBtnImitation.OffsetImage True, 0, 0 ' reset image
        ' we'll use another property to see if the mouse is over our control when the up event occurs
        If aicBtnImitation.isMouseOver = False Then aicBtnImitation.IntensityOffset = -5
    End If
End Sub

Private Sub aicBtnImitation_MouseEnter()
    aicBtnImitation.IntensityOffset = 0 ' make image appear slightly lighter
End Sub

Private Sub aicBtnImitation_MouseExit()
    aicBtnImitation.IntensityOffset = -5 ' make image appear slightly darker
End Sub

Private Sub Form_Load()
    ' in design time, I want the bubble transparent, but then I can't see it
    ' so it is set to 100% opacity and here, I set it to zero
    aicBubble.Opacity = 0&
        
    ' size our ole drop control
    ucDragDrop.Move Shape1.Left + 1, Shape1.Top + 1, Shape1.Width - 2, Shape1.Height - 2
    
    ' set the combo listindex
    cboPlayTime.ListIndex = 0
    
End Sub

Private Sub Form_Resize()
    If Me.WindowState = vbMinimized Then
        ' don't rotate clock hands while we are minimized
        Timer1.Enabled = False
    Else
        ' if coming out of minimize state, start up clock again
        If Timer1.Enabled = False Then InitializeClock
    End If
    
End Sub

Private Sub mnuSample_Click(Index As Integer)
    If Index = 0 Then ' load the res image
        Call ucDragDrop_DblClick(vbLeftButton)
    ElseIf Index = 2 Then ' clear image
        ucDragDrop.ClearImage
    End If
End Sub

Private Sub Timer1_Timer()
    Dim tTime As Date
    tTime = Time
    If Second(tTime) = 0 Then
        ' update all three hands: hour, minute, second; else just the second hand
        aicHour.Rotation() = 30 * Hour(tTime) + (Minute(tTime) / 60) * 24
        aicMinute.Rotation() = 6 * Minute(tTime)
    End If
    aicSecond.Rotation() = 6 * Second(tTime)
End Sub

Private Sub InitializeClock()
    Dim tTime As Date
    ' update all three hands: hour, minute, second
    tTime = Time
    aicHour.Rotation() = 30 * Hour(tTime) + (Minute(tTime) / 60) * 24
    aicMinute.Rotation() = 6 * Minute(tTime)
    aicSecond.Rotation() = 6 * Second(tTime)
    ' ensure timer interval set & enabled
    Timer1.Interval = 1000
    Timer1.Enabled = True
End Sub

Private Sub aicCheetah_Click(ByVal Button As Integer)
    Unload Me   ' testing click event
End Sub

Private Sub aicCheetah_MouseEnter()
    ' on mouse enter, we will grayscale and fade the bubble in
    aicCheetah.ShadowEnabled = False
    aicCheetah.grayScale = aiNoGrayScale
    aicBubble.FadeInOut 100
End Sub

Private Sub aicCheetah_MouseExit()
    ' on mouse exit we will fade bubble out, then when it
    ' is completely faded out, we will grayscale cheetah
    aicCheetah.ShadowEnabled = True
    aicBubble.FadeInOut 0
End Sub

Private Sub aicBubble_FadeTerminated(ByVal CurrentOpacity As Long)
    ' grayscale cheetah when bubble is faded out
    If CurrentOpacity = 0& Then aicCheetah.grayScale = aiRedGreenMask
End Sub

Private Sub ucDragDrop_Click(ByVal Button As Integer)
    ' show a popup menu when right clicked upon
    If Button = vbRightButton Then PopupMenu mnuPopup, , , , mnuSample(0)
End Sub

Private Sub ucDragDrop_DblClick(ByVal Button As Integer)
    If Button = vbLeftButton Then
        ' load a res file image when double clicked upon
        ucDragDrop.LoadImage_FromResource VB.Global, "Custom", "LaVolpe"
    End If
End Sub

Private Sub ucDragDrop_OLEDragDrop(Data As DataObject, Effect As Long, Button As Integer, Shift As Integer, x As Single, y As Single)
    ' simple example showing control can receive ole drop operations
    If Data.Files.Count > 0 Then
        If ucDragDrop.LoadImage_FromFile(Data.Files(1), ucDragDrop.Width, ucDragDrop.Height) = False Then
            MsgBox "Could not load that file", vbInformation + vbOKOnly
        End If
    End If
End Sub

Private Sub cboPlayTime_Click()
    If Me.Visible = True Then
    
        Select Case cboPlayTime.ListIndex
        Case 0 ' Restore
            ucPlayTime.InversedImage = False
            ucPlayTime.grayScale = aiNoGrayScale
            ucPlayTime.Mirror = aiMirrorNone
            ucPlayTime.Rotation = 0&
            ' to use the following function, the KeepOriginalBytes property must = True
            ucPlayTime.LoadImage_FromOrignalBytes
        Case 1 'Invert Colors
            ucPlayTime.InversedImage = Not ucPlayTime.InversedImage
        Case 2 ' Mirror Horizontal
            ucPlayTime.Mirror = ucPlayTime.Mirror Xor aiMirrorHorizontal
        Case 3 'Shift Pixels
        
            ' this is a simple example, showing that you can extract the
            ' image's 32bpp bytes and modify them to your heart's desire
            ' the send them back to the control.
            
            Dim x As Long, y As Long, dibBytes() As Byte
            Dim cX As Long
            Dim xOffset As Long, lSwap As Long
            
            ' get the bytes into a 2D array. There are other options available
            If ucPlayTime.GetImageBytes(dibBytes, 0&, True, , , True) = True Then
                cX = UBound(dibBytes, 1) \ 4 + 1    ' width of image
                For y = 0 To UBound(dibBytes, 2)
                    xOffset = UBound(dibBytes, 1) - 3   ' position of last pixel
                    For x = 0 To (cX \ 3) * 4 - 4 Step 4    ' do 1/3 of image
                        ' we will swap the last 1/3 with the first 1/3
                        CopyMemory lSwap, dibBytes(xOffset, y), 4&
                        CopyMemory dibBytes(xOffset, y), dibBytes(x, y), 4&
                        CopyMemory dibBytes(x, y), lSwap, 4&
                        xOffset = xOffset - 4
                    Next
                Next
                ' set the bytes, setting the correct parameters
                Call ucPlayTime.SetImageBytes(dibBytes)
            End If
            
        Case 4 ' Rotate 180
            ' note: if you expect your image to be rotated you should
            ' set AutoResize=False and ScaleMethod=aiLockScale after you size your control
            ' Of course, set the RotateEnabled property to True also
            ucPlayTime.Rotation = ucPlayTime.Rotation + 180
            
        Case 5: ' grayscale
            ucPlayTime.grayScale = ucPlayTime.grayScale Xor aiNTSCPAL
            
        Case 6 ' Fade
            ucPlayTime.FadeInOut 0, 5, 70
        End Select
        
    End If
End Sub

Private Sub ucPlayTime_FadeTerminated(ByVal CurrentOpacity As Long)
    ' our playtime image - fade back in when faded out
    If CurrentOpacity = 0 Then ucPlayTime.FadeInOut 100, 5, 70
End Sub
