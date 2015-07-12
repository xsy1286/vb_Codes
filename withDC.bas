Attribute VB_Name = "withDC"
Option Explicit
Private Type POINTAPI
 x As Long
 y As Long
End Type
Public tmpP As POINTAPI
'way of use:BitBlt Picdown.hdc, a, b, c, d, GetDC(0), 0, 0, vbSrcCopy
Public Declare Function BitBlt Lib "gdi32" (ByVal hDestDC As Long, ByVal x As Long, ByVal y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal dwRop As Long) As Long
'在NT环境下，如在一次世界传输中要求在源设备场景中进行剪切或旋转处理，这个函数的执行会失败
'如目标和源DC的映射关系要求矩形中像素的大小必须在传输过程中改变，那么这个函数会根据需要自动伸缩、旋转、折叠、或切断，以便完成最终的传输过程
Public Declare Function StretchBlt Lib "gdi32" (ByVal hdc As Long, ByVal x As Long, ByVal y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal nSrcWidth As Long, ByVal nSrcHeight As Long, ByVal dwRop As Long) As Long
'可用GetDeviceCaps函数判断特定的设备场景是否支持此函数
'不可选择对源位图进行剪切或旋转处理，源位图也不能是一个图元文件设备场景
Public Declare Function GetDeviceCaps Lib "gdi32" (ByVal hdc As Long, ByVal nIndex As Long) As Long
'StretchBlt与BitBlt不同在于StretchBlt方法能够延伸或收缩位图以适应目标区域的大小

Public Declare Function GetDC Lib "user32" (ByVal hwnd As Long) As Long
Public Declare Function ReleaseDC Lib "user32" (ByVal hwnd As Long, ByVal hdc As Long) As Long
Public Declare Function GetPixel Lib "gdi32" (ByVal hdc As Long, ByVal x As Long, ByVal y As Long) As Long
Public Const Srccopy = &HCC0020


Public Declare Function MoveToEx Lib "gdi32" (ByVal hdc As Long, ByVal x As Long, ByVal y As Long, lpPoint As POINTAPI) As Long
Public Declare Function LineTo Lib "gdi32" (ByVal hdc As Long, ByVal x As Long, ByVal y As Long) As Long
