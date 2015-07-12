VERSION 5.00
Begin VB.Form Form1 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "改变屏幕分辨率"
   ClientHeight    =   1680
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   3060
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   1680
   ScaleWidth      =   3060
   StartUpPosition =   3  '窗口缺省
   Begin VB.Frame Frame1 
      Caption         =   "选择分辨率模式"
      Height          =   990
      Left            =   150
      TabIndex        =   1
      Top             =   90
      Width           =   2730
      Begin VB.ComboBox Combo1 
         Height          =   300
         Left            =   195
         Style           =   2  'Dropdown List
         TabIndex        =   2
         Top             =   345
         Width           =   2340
      End
   End
   Begin VB.CommandButton Command1 
      Caption         =   "立即改变"
      Height          =   345
      Left            =   930
      TabIndex        =   0
      Top             =   1215
      Width           =   990
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'download by http://www.codefans.net
Private Declare Function lstrcpy Lib "kernel32" Alias "lstrcpyA" (lpString1 As Any, lpString2 As Any) As Long
Private Const CCHDEVICENAME = 32
Private Const CCHFORMNAME = 32
Private Const ENUM_CURRENT_SETTINGS = 1
Private Type DEVMODE
        dmDeviceName As String * CCHDEVICENAME
        dmSpecVersion As Integer
        dmDriverVersion As Integer
        dmSize As Integer
        dmDriverExtra As Integer
        dmFields As Long
        dmOrientation As Integer
        dmPaperSize As Integer
        dmPaperLength As Integer
        dmPaperWidth As Integer
        dmScale As Integer
        dmCopies As Integer
        dmDefaultSource As Integer
        dmPrintQuality As Integer
        dmColor As Integer
        dmDuplex As Integer
        dmYResolution As Integer
        dmTTOption As Integer
        dmCollate As Integer
        dmFormName As String * CCHFORMNAME
        dmUnusedPadding As Integer
        dmBitsPerPel As Long
        dmPelsWidth As Long
        dmPelsHeight As Long
        dmDisplayFlags As Long
        dmDisplayFrequency As Long
End Type

Private Declare Function ChangeDisplaySettings Lib "user32" Alias "ChangeDisplaySettingsA" (ByVal lpDevMode As Long, ByVal dwflags As Long) As Long
Private Declare Function EnumDisplaySettings Lib "user32" Alias "EnumDisplaySettingsA" (ByVal lpszDeviceName As String, ByVal iModeNum As Long, lpDevMode As Any) As Long
Private Declare Function GetSystemMetrics Lib "user32" (ByVal nIndex As Long) As Long

Private Const SM_CXSCREEN = 0
Private Const SM_CYSCREEN = 1

Dim pNewMode As DEVMODE
Dim pOldMode As Long
Dim nOrgWidth As Integer, nOrgHeight As Integer
    
'设置显示器分辨率的执行函数
Private Function SetDisplayMode(Width As Integer, Height As Integer, Color As Integer) As Long ', Freq As Long) As Long
    On Error GoTo ErrorHandler
    Const DM_PELSWIDTH = &H80000
    Const DM_PELSHEIGHT = &H100000
    Const DM_BITSPERPEL = &H40000
    Const DM_DISPLAYFLAGS = &H200000
    Const DM_DISPLAYFREQUENCY = &H400000
    With pNewMode
        .dmSize = Len(pNewMode)
        If Color = 0 Then 'Color = 0 时不更改屏幕颜色
            .dmFields = DM_PELSWIDTH Or DM_PELSHEIGHT
        Else
            .dmFields = DM_BITSPERPEL Or DM_PELSWIDTH Or DM_PELSHEIGHT  'Or DM_DISPLAYFREQUENCY'属性率的更改还是没办法,不过,不加入此DM_DISPLAYFREQUENCY这个参数,只要系统支持,应该不会更改刷新率的
        End If
        .dmPelsWidth = Width
        .dmPelsHeight = Height
        If Color <> 0 Then
        .dmBitsPerPel = Color
        End If
    End With
    pOldMode = lstrcpy(pNewMode, pNewMode)
    SetDisplayMode = ChangeDisplaySettings(pOldMode, 1)
    Exit Function
ErrorHandler:
    MsgBox Err.Description, vbCritical, "VB广场"
End Function

Private Sub Command1_Click()
    Dim nWidth As Integer, nHeight As Integer, nColor As Integer
    Select Case Combo1.ListIndex
        Case 0
            nWidth = 640: nHeight = 480: nColor = 16  '640*480*16位真彩色,256色nColor = 8,16色nColor = 4,nColor = 0 表示不改变颜色
        Case 1
            nWidth = 640: nHeight = 480: nColor = 24
        Case 2
            nWidth = 640: nHeight = 480: nColor = 32
        Case 3
            nWidth = 800: nHeight = 600: nColor = 16
        Case 4
            nWidth = 800: nHeight = 600: nColor = 24
        Case 5
            nWidth = 800: nHeight = 600: nColor = 32
        Case 6
            nWidth = 1024: nHeight = 768: nColor = 16
        Case 7
            nWidth = 1024: nHeight = 768: nColor = 24
        Case 8
            nWidth = 1024: nHeight = 768: nColor = 32
        Case other
            nWidth = 800: nHeight = 600: nColor = 16
    End Select
    Call SetDisplayMode(nWidth, nHeight, nColor)  '注意,系统不支持的显示模式不能选,否则,准备用安全模式重启动吧.API函数EnumDisplaySettings可以选择系统支持的模式,自己去写吧,也很简单.如果你还有什么问题,请给我发信或留言.
End Sub

Private Sub Form_Load()
    Combo1.AddItem "640*480*16位真彩色"
    Combo1.AddItem "640*480*24位真彩色"
    Combo1.AddItem "640*480*32位真彩色"
    Combo1.AddItem "800*600*16位真彩色"
    Combo1.AddItem "800*600*24位真彩色"
    Combo1.AddItem "800*600*32位真彩色"
    Combo1.AddItem "1024*768*16位真彩色"
    Combo1.AddItem "1024*768*24位真彩色"
    Combo1.AddItem "1024*768*32位真彩色"
    Combo1.Text = Combo1.List(0)
    nOrgWidth = GetDisplayWidth
    nOrgHeight = GetDisplayHeight
    'nOrgWidth = GetSystemMetrics(SM_CXSCREEN)'两种获取初始屏幕大小的方法均可
    'nOrgHeight = GetSystemMetrics(SM_CYSCREEN)
End Sub

Private Function GetDisplayWidth() As Integer
    GetDisplayWidth = Screen.Width \ Screen.TwipsPerPixelX
End Function

Private Function GetDisplayHeight() As Integer
    GetDisplayHeight = Screen.Height \ Screen.TwipsPerPixelY
End Function

Private Sub RestoreDisplayMode()
    Call SetDisplayMode(nOrgWidth, nOrgHeight, 0)
End Sub

Private Sub Form_Unload(Cancel As Integer)
    RestoreDisplayMode
End Sub
