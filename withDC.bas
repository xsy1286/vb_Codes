Attribute VB_Name = "withDC"
Option Explicit
Private Type POINTAPI
 x As Long
 y As Long
End Type
Public tmpP As POINTAPI
'way of use:BitBlt Picdown.hdc, a, b, c, d, GetDC(0), 0, 0, vbSrcCopy
Public Declare Function BitBlt Lib "gdi32" (ByVal hDestDC As Long, ByVal x As Long, ByVal y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal dwRop As Long) As Long
'��NT�����£�����һ�����紫����Ҫ����Դ�豸�����н��м��л���ת�������������ִ�л�ʧ��
'��Ŀ���ԴDC��ӳ���ϵҪ����������صĴ�С�����ڴ�������иı䣬��ô��������������Ҫ�Զ���������ת���۵������жϣ��Ա�������յĴ������
Public Declare Function StretchBlt Lib "gdi32" (ByVal hdc As Long, ByVal x As Long, ByVal y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal nSrcWidth As Long, ByVal nSrcHeight As Long, ByVal dwRop As Long) As Long
'����GetDeviceCaps�����ж��ض����豸�����Ƿ�֧�ִ˺���
'����ѡ���Դλͼ���м��л���ת����ԴλͼҲ������һ��ͼԪ�ļ��豸����
Public Declare Function GetDeviceCaps Lib "gdi32" (ByVal hdc As Long, ByVal nIndex As Long) As Long
'StretchBlt��BitBlt��ͬ����StretchBlt�����ܹ����������λͼ����ӦĿ������Ĵ�С

Public Declare Function GetDC Lib "user32" (ByVal hwnd As Long) As Long
Public Declare Function ReleaseDC Lib "user32" (ByVal hwnd As Long, ByVal hdc As Long) As Long
Public Declare Function GetPixel Lib "gdi32" (ByVal hdc As Long, ByVal x As Long, ByVal y As Long) As Long
Public Const Srccopy = &HCC0020


Public Declare Function MoveToEx Lib "gdi32" (ByVal hdc As Long, ByVal x As Long, ByVal y As Long, lpPoint As POINTAPI) As Long
Public Declare Function LineTo Lib "gdi32" (ByVal hdc As Long, ByVal x As Long, ByVal y As Long) As Long
