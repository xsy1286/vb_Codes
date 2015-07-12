Attribute VB_Name = "MScreen"

Option Explicit
'******************************
'  源码学习下载www.lvcode.com
'    欢迎分享源码给Love代码
'******************************
Public Declare Function CreateCompatibleDC Lib "GDI32.DLL" (ByVal hDC As Long) As Long
Public Declare Function CreateCompatibleBitmap Lib "GDI32.DLL" (ByVal hDC As Long, ByVal nWidth As Long, ByVal nHeight As Long) As Long
Public Declare Function GetDeviceCaps Lib "GDI32.DLL" (ByVal hDC As Long, ByVal iCapabilitiy As Long) As Long
Public Declare Function GetSystemPaletteEntries Lib "GDI32.DLL" (ByVal hDC As Long, ByVal wStartIndex As Long, ByVal wNumEntries As Long, lpPaletteEntries As PALETTEENTRY) As Long
Public Declare Function CreatePalette Lib "GDI32.DLL" (lpLogPalette As LOGPALETTE) As Long
Public Declare Function SelectObject Lib "GDI32.DLL" (ByVal hDC As Long, ByVal hObject As Long) As Long
Public Declare Function BitBlt Lib "GDI32.DLL" (ByVal hDCDest As Long, ByVal XDest As Long, ByVal YDest As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal HDCSRC As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal dwRop As Long) As Long
Public Declare Function DeleteDC Lib "GDI32.DLL" (ByVal hDC As Long) As Long
Public Declare Function GetForegroundWindow Lib "USER32.DLL" () As Long
Public Declare Function SelectPalette Lib "GDI32.DLL" (ByVal hDC As Long, ByVal hPalette As Long, ByVal bForceBackground As Long) As Long
Public Declare Function RealizePalette Lib "GDI32.DLL" (ByVal hDC As Long) As Long
Public Declare Function GetWindowDC Lib "USER32.DLL" (ByVal hWnd As Long) As Long
Public Declare Function GetDC Lib "USER32.DLL" (ByVal hWnd As Long) As Long
Public Declare Function GetWindowRect Lib "USER32.DLL" (ByVal hWnd As Long, lpRect As RECT) As Long
Public Declare Function ReleaseDC Lib "USER32.DLL" (ByVal hWnd As Long, ByVal hDC As Long) As Long
Public Declare Function GetDesktopWindow Lib "USER32.DLL" () As Long
Public Declare Function OleCreatePictureIndirect Lib "olepro32.dll" (PicDesc As PicBmp, RefIID As GUID, ByVal fPictureOwnsHandle As Long, IPic As IPicture) As Long

Public Type PicBmp
    Size As Long
    Type As Long
    HBMP As Long
    HPal As Long
    Reserved As Long
End Type

Public Type PALETTEENTRY
    PERed As Byte
    PEGreen As Byte
    PEBlue As Byte
    PEFlags As Byte
End Type

Public Type LOGPALETTE
    PALVersion As Integer
    PALNumEntries As Integer
    PALPalEntry(255) As PALETTEENTRY
End Type

Public Type GUID
    Data1 As Long
    Data2 As Integer
    Data3 As Integer
    Data4(7) As Byte
End Type

Public Type RECT
    Left As Long
    Top As Long
    Right As Long
    Bottom As Long
End Type

Public Const RASTERCAPS As Long = 38
Public Const RC_PALETTE As Long = &H100
Public Const SIZEPALETTE As Long = 104

Public Function CreateBitmapPicture(ByVal HBMP As Long, ByVal HPal As Long) As Picture
  
    Dim Pic As PicBmp
    Dim IPic As IPicture
    Dim IID_IDispatch As GUID

    On Error Resume Next
    
    With IID_IDispatch
        .Data1 = &H20400
        .Data4(0) = &HC0
        .Data4(7) = &H46
    End With
    With Pic
        .Size = Len(Pic)
        .Type = vbPicTypeBitmap
        .HBMP = HBMP
        .HPal = HPal
    End With

    OleCreatePictureIndirect Pic, IID_IDispatch, 1, IPic
    Set CreateBitmapPicture = IPic
  
End Function

Public Function CaptureWindow(ByVal HWNDSrc As Long, ByVal Client As Boolean, ByVal LeftSRC As Long, ByVal TopSRC As Long, ByVal WidthSRC As Long, ByVal HeightSRC As Long) As Picture
  
    Dim HDCMemory As Long
    Dim HBMP As Long
    Dim HBMPPrev As Long
    Dim HDCSRC As Long
    Dim HPal As Long
    Dim HPalPrev As Long
    Dim RasterCapsScrn As Long
    Dim HasPaletteScrn As Long
    Dim PaletteSizeScrn As Long
    Dim LogPal As LOGPALETTE
  
    On Error Resume Next
  
    If Client Then
        HDCSRC = GetDC(HWNDSrc)
    Else
        HDCSRC = GetWindowDC(HWNDSrc)
    End If
    HDCMemory = CreateCompatibleDC(HDCSRC)
    HBMP = CreateCompatibleBitmap(HDCSRC, WidthSRC, HeightSRC)
    HBMPPrev = SelectObject(HDCMemory, HBMP)
    RasterCapsScrn = GetDeviceCaps(HDCSRC, RASTERCAPS)
    HasPaletteScrn = RasterCapsScrn And RC_PALETTE
    PaletteSizeScrn = GetDeviceCaps(HDCSRC, SIZEPALETTE)
    If HasPaletteScrn And (PaletteSizeScrn = 256) Then
        LogPal.PALVersion = &H300
        LogPal.PALNumEntries = 256
        GetSystemPaletteEntries HDCSRC, 0, 256, LogPal.PALPalEntry(0)
        HPal = CreatePalette(LogPal)
        HPalPrev = SelectPalette(HDCMemory, HPal, 0)
        RealizePalette HDCMemory
    End If
    BitBlt HDCMemory, 0, 0, WidthSRC, HeightSRC, HDCSRC, LeftSRC, TopSRC, vbSrcCopy
    HBMP = SelectObject(HDCMemory, HBMPPrev)
    If HasPaletteScrn And (PaletteSizeScrn = 256) Then
        HPal = SelectPalette(HDCMemory, HPalPrev, 0)
    End If
    DeleteDC HDCMemory
    ReleaseDC HWNDSrc, HDCSRC
    Set CaptureWindow = CreateBitmapPicture(HBMP, HPal)
   
End Function

Public Function CaptureScreen() As Picture
  
    Dim HWNDScreen As Long
  
    On Error Resume Next

    HWNDScreen = GetDesktopWindow()
    Set CaptureScreen = CaptureWindow(HWNDScreen, False, 0, 0, Screen.Width \ Screen.TwipsPerPixelX, Screen.Height \ Screen.TwipsPerPixelY)
  
End Function

Public Function TakeScreenShot(PictureBox As Form, ByVal FileOutput As String)

    PictureBox.Picture = CaptureScreen
    
End Function
