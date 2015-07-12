VERSION 5.00
Begin VB.UserControl pngTransparent 
   ClientHeight    =   3600
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   4800
   ScaleHeight     =   3600
   ScaleWidth      =   4800
   Begin VB.Timer Timer2 
      Left            =   1560
      Top             =   1920
   End
End
Attribute VB_Name = "pngTransparent"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Public Event MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
Public Event MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
Public Event MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
'****PNG
Private Declare Function GdiplusStartup Lib "gdiplus" (token As Long, inputbuf As GdiplusStartupInput, Optional ByVal outputbuf As Long = 0) As GpStatus
Private Declare Sub GdiplusShutdown Lib "gdiplus" (ByVal token As Long)

Private Declare Function GdipCreateFromHDC Lib "gdiplus" (ByVal hwnd As Long, graphics As Long) As GpStatus
Private Declare Function GetDC Lib "user32" (ByVal hwnd As Long) As Long

Private Declare Function GdipDeleteGraphics Lib "gdiplus" (ByVal graphics As Long) As GpStatus
Private Declare Function GdipDrawImageRect Lib "gdiplus" (ByVal graphics As Long, ByVal image As Long, ByVal x As Single, ByVal y As Single, ByVal Width As Single, ByVal Height As Single) As GpStatus
Private Declare Function GdipLoadImageFromFile Lib "gdiplus" (ByVal fileName As String, image As Long) As GpStatus
Private Declare Function GdipGetImageWidth Lib "gdiplus" (ByVal image As Long, Width As Long) As GpStatus
Private Declare Function GdipGetImageHeight Lib "gdiplus" (ByVal image As Long, Height As Long) As GpStatus
Private Declare Function GdipDisposeImage Lib "gdiplus" (ByVal image As Long) As GpStatus

Private Type GdiplusStartupInput
    GdiplusVersion As Long
    DebugEventCallback As Long
    SuppressBackgroundThread As Long
    SuppressExternalCodecs As Long
End Type

Private Enum GpStatus
    Ok = 0
    GenericError = 1
    InvalidParameter = 2
    OutOfMemory = 3
    ObjectBusy = 4
    InsufficientBuffer = 5
    NotImplemented = 6
    Win32Error = 7
    WrongState = 8
    Aborted = 9
    FileNotFound = 10
    ValueOverflow = 11
    AccessDenied = 12
    UnknownImageFormat = 13
    FontFamilyNotFound = 14
    FontStyleNotFound = 15
    NotTrueTypeFont = 16
    UnsupportedGdiplusVersion = 17
    GdiplusNotInitialized = 18
    PropertyNotFound = 19
    PropertyNotSupported = 20
End Enum
Dim m_token As Long
Private urlpng As String    '要显示的图片名称和路径。
'***** PNG

'**** 透明
Private Declare Function GetWindowLong Lib "user32" Alias "GetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long) As Long
Private Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
Private Declare Function SetLayeredWindowAttributes Lib "user32" (ByVal hwnd As Long, ByVal crKey As Long, ByVal bAlpha As Byte, ByVal dwFlags As Long) As Long
Const WS_EX_LAYERED = &H80000
Const GWL_EXSTYLE = (-20)
Const LWA_COLORKEY = &H1
'******透明

'****移动

'****移动

Public Enum backstyenum
透明 = 0

自设 = 1
End Enum
Dim backstyeg As backstyenum
Dim bcolor(0 To 2) As Integer


Public Function bc(a As Integer, b As Integer, c As Integer) As Long
bcolor(0) = a: bcolor(1) = b: bcolor(2) = c
If backstyeg = 2 Then UserControl.BackColor = RGB(a, b, c)
End Function

Private Sub UserControl_Initialize()
'
Dim transcolor   As Long
   Dim StartupInput As GdiplusStartupInput
     StartupInput.GdiplusVersion = 1
     If GdiplusStartup(m_token, StartupInput, ByVal 0) Then
             MsgBox "Error initializing GDI+"
             Exit Sub
     End If
   
   Timer2.Interval = 20
   
Select Case backstyeg
Case 0
 transcolor = RGB(66, 66, 66) '必须都66
UserControl.BackColor = transcolor
Case 1
UserControl.BackColor = RGB(bcolor(0), bcolor(1), bcolor(2))
End Select

End Sub

Private Sub UserControl_Paint()
    Dim pImg As Long
     Dim pGraphics As Long
     Dim w As Long, h As Long
    
     Call GdipCreateFromHDC(GetDC(UserControl.hwnd), pGraphics)
     Call GdipLoadImageFromFile(StrConv(urlpng, vbUnicode), pImg)
     Call GdipGetImageWidth(pImg, w)
     Call GdipGetImageHeight(pImg, h)
     Call GdipDrawImageRect(pGraphics, pImg, 0, 0, w, h)
     
     Call GdipDisposeImage(pImg)
     Call GdipDeleteGraphics(pGraphics)
    
End Sub
Public Property Let url(ByVal urlin As String)
   urlpng = urlin
    PropertyChanged "url"
        Dim pImg As Long
     Dim pGraphics As Long
     Dim w As Long, h As Long
    
     Call GdipCreateFromHDC(GetDC(UserControl.hwnd), pGraphics)
     Call GdipLoadImageFromFile(StrConv(urlpng, vbUnicode), pImg)
     Call GdipGetImageWidth(pImg, w)
     Call GdipGetImageHeight(pImg, h)
     Call GdipDrawImageRect(pGraphics, pImg, 0, 0, w, h)
     
     Call GdipDisposeImage(pImg)
     Call GdipDeleteGraphics(pGraphics)
End Property
Public Property Get url() As String
    url = urlpng
End Property
Private Sub UserControl_ReadProperties(PropBag As PropertyBag)
    With PropBag
     '  url = .ReadProperty("Url", urlpng)
     
    End With
   
End Sub

Private Sub UserControl_WriteProperties(PropBag As PropertyBag)
    With PropBag
       ' .WriteProperty "Url", urlpng
     
    End With
End Sub
Public Property Get MouseDown() As String
    url = urlpng
End Property



Public Property Get backsty() As backstyenum
backsty = backstyeg
End Property
Public Property Let backsty(ByVal NewValue As backstyenum)
backstyeg = NewValue
PropertyChanged "BackSty"
   If backstyeg = 2 Then
UserControl.BackColor = RGB(bcolor(0), bcolor(1), bcolor(2))
End If

End Property

Private Sub UserControl_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
    RaiseEvent MouseDown(Button, Shift, Int(ScaleX(x, vbPixels, vbContainerPosition)), Int(ScaleY(y, vbPixels, vbContainerPosition)))
End Sub
Private Sub UserControl_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    RaiseEvent MouseMove(Button, Shift, Int(ScaleX(x, vbPixels, vbContainerPosition)), Int(ScaleY(y, vbPixels, vbContainerPosition)))
End Sub

Private Sub UserControl_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
    RaiseEvent MouseUp(Button, Shift, Int(ScaleX(x, vbPixels, vbContainerPosition)), Int(ScaleY(y, vbPixels, vbContainerPosition)))
End Sub
