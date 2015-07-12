Attribute VB_Name = "Module1"
   'module 需注意Form1的物件需要Form dot
Option Explicit
Public Declare Function PrintWindow Lib "user32" (ByVal hwnd As Long, ByVal hdcBlt As Long, ByVal nFlags As Long) As Long
Public Declare Function BitBlt Lib "gdi32" (ByVal hDestDC As Long, ByVal x As Long, ByVal y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal dwRop As Long) As Long
Public Declare Function StretchBlt Lib "gdi32" (ByVal hdc As Long, ByVal x As Long, ByVal y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal nSrcWidth As Long, ByVal nSrcHeight As Long, ByVal dwRop As Long) As Long
Public Declare Function GetDC Lib "user32" (ByVal hwnd As Long) As Long
Public Declare Function GetCursorPos Lib "user32" (lpPonit As POINTAPI) As Long

Public Declare Function SetWindowsHookEx Lib "user32" Alias "SetWindowsHookExA" (ByVal idHook As Long, ByVal lpfn As Long, ByVal hmod As Long, ByVal dwThreadId As Long) As Long
Public Declare Function UnhookWindowsHookEx Lib "user32" (ByVal hHook As Long) As Long
Public Declare Function CallNextHookEx Lib "user32" (ByVal hHook As Long, ByVal nCode As Long, ByVal wParam As Long, lParam As Any) As Long
Private Declare Sub CopyMemory Lib "kernel32" _
                          Alias "RtlMoveMemory" _
                          (Destination As Any, _
                          Source As Any, _
                          ByVal Length As Long)
Public l, h As Long
Public tt As Integer
Public bb As Boolean
Public d, dd As Integer
Public Type POINTAPI
    x As Long
    y As Long
End Type
Private ret As Long

Private Type MSLLHOOKSTRUCT  'module中API不能private，但Type可以
    pt As POINTAPI
    mouseData As Long
    Flags As Long
    time As Long
    dwExtraInfo As Long
End Type

Public Const Srccopy = &HCC0020

Public Const WH_KEYBOARD_LL = 13
Public Const WH_MOUSE_LL = 14
Public t As Long
'消息
Public Const HC_ACTION = 0
Public Const HC_SYSMODALOFF = 5
Public Const HC_SYSMODALON = 4
'鼠标消息
Public Const WM_MOUSEMOVE = &H200
Public Const WM_LBUTTONDOWN = &H201
Public Const WM_LBUTTONUP = &H202
Public Const WM_LBUTTONDBLCLK = &H203
Public Const WM_RBUTTONDOWN = &H204
Public Const WM_RBUTTONUP = &H205
Public Const WM_RBUTTONDBLCLK = &H206
Public Const WM_MBUTTONDOWN = &H207
Public Const WM_MBUTTONUP = &H208
Public Const WM_MBUTTONDBLCLK = &H209
Public Const WM_MOUSEACTIVATE = &H21
Public Const WM_MOUSEFIRST = &H200
Public Const WM_MOUSELAST = &H209
Public Const WM_MOUSEWHEEL = &H20A



Public pt As POINTAPI, pt2 As POINTAPI
Public lHook As Long
Public hHook As Long

Public Function MouseHook(ByVal nCode As Long, _
                       ByVal wParam As Long, _
                       ByVal lParam As Long) As Long

    Dim mhs As MSLLHOOKSTRUCT

    If wParam = WM_LBUTTONDOWN And d = 0 Then
    d = 1
  
        Call CopyMemory(mhs, ByVal lParam, LenB(mhs))
        pt = mhs.pt
'    Call ShapeDraw
  
     Form2.Shape1.Visible = True

      '  Debug.Print "左键单击    坐标:" & pt.x & "  "; pt.y
    End If
    
    
    
     If wParam = WM_LBUTTONUP And d = 1 Then
          Call CopyMemory(mhs, ByVal lParam, LenB(mhs))
        pt2 = mhs.pt
     d = 0
    
        'Debug.Print "左键起    坐标:" & pt2.x & "  "; pt2.y
        
           Call ShapeDraw
       
      
        If pt.x <> pt2.x Or pt2.y <> pt.y Then
'先false 再true防止显示系统性错乱
 tt = 1
Call Draw
End If

End If


If wParam = WM_MOUSEMOVE And d = 1 Then
        Call CopyMemory(mhs, ByVal lParam, LenB(mhs))     '本例使用Len 函数返回字符串的总字符数。 　　Dim n As Integer 　　n=LenB("Hello world") //返回11
        If dn = 1 Then dn = 0
        pt2 = mhs.pt
       
           Call ShapeDraw
    End If
    Call CallNextHookEx(hHook, nCode, wParam, lParam)
End Function
Public Sub ShapeDraw()
If pt2.x > pt.x Then
      If pt2.y > pt.y Then
      Form2.Shape1.Left = pt.x - 2: Form2.Shape1.Top = pt.y - 2
      'Form2.Command1.Caption = Str(pt.X): Form2.Command1.Caption = Str(pt.Y)
Form2.Shape1.Width = (pt2.x - pt.x) + 2

Form2.Shape1.Height = (pt2.y - pt.y) + 2

Else:
Form2.Shape1.Left = pt.x - 2: Form2.Shape1.Top = pt2.y - 2
Form2.Shape1.Width = (pt2.x - pt.x) + 2
Form2.Shape1.Height = (pt.y - pt2.y) + 2

End If

End If


      If pt2.x < pt.x Then
      If pt2.y > pt.y Then
      Form2.Shape1.Left = pt2.x - 2: Form2.Shape1.Top = pt.y - 2
Form2.Shape1.Width = (pt.x - pt2.x) + 2
Form2.Shape1.Height = (pt2.y - pt.y) + 2

Else:

Form2.Shape1.Width = (pt.x - pt2.x) + 2
Form2.Shape1.Height = (pt.y - pt2.y) + 2
Form2.Shape1.Left = pt2.x - 2: Form2.Shape1.Top = pt2.y - 2
End If

End If
End Sub

Public Sub Draw()
Form1.Picture1.Cls '注意清除后重绘
     Form2.Command3.Visible = False
      Form2.Command4.Visible = False


Call delay(150)
      
      'module 需注意Form1的物件需要Form dot
      If pt2.x > pt.x Then
      
       If pt2.y > pt.y Then

Form1.Picture1.Width = (pt2.x - pt.x)

Form1.Picture1.Height = (pt2.y - pt.y)
l = (pt2.x - pt.x)
h = (pt2.y - pt.y)
Form2.Label2.Caption = Str(l) + "   " + Str(h)
                            '此为像素点距离
                               '↓    ↓
                            
 BitBlt Form1.Picture1.hdc, (0 - pt.x), (0 - pt.y), Form1.Picture1.Width + 1300, Form1.Picture1.Height + 1300, GetDC(0), 0, 0, vbSrcCopy
'Debug.Print "r=" & Str(0 - pt.X) & Str(0 - pt.Y)
'Form1.Picture1.PaintPicture Form2.Picture, 0, 0, Form1.Picture1.Width, Form1.Picture1.Height, pt.X  , pt.Y  , Screen.Width, Screen.Height, vbSrcCopy
  Form2.Command3.Left = pt2.x - 66.67: Form2.Command3.Top = pt2.y - 20: Form2.Command4.Left = pt2.x - 33.34: Form2.Command4.Top = pt2.y - 20
dn = 1
Else:
Form1.Picture1.Width = (pt2.x - pt.x)
Form1.Picture1.Height = (pt.y - pt2.y)

BitBlt Form1.Picture1.hdc, 0 - pt.x, 0 - pt2.y, Form1.Picture1.Width + 1300, Form1.Picture1.Height + 1300, GetDC(0), 0, 0, vbSrcCopy
Form2.Command3.Left = pt2.x - 66.67: Form2.Command3.Top = pt2.y + 20: Form2.Command4.Left = pt2.x - 33.34: Form2.Command4.Top = pt2.y + 20
dn = 1
End If

End If


      If pt2.x < pt.x Then
      If pt2.y > pt.y Then
Form1.Picture1.Width = (pt.x - pt2.x)
Form1.Picture1.Height = (pt2.y - pt.y)

BitBlt Form1.Picture1.hdc, 0 - pt2.x, 0 - pt.y, Form1.Picture1.Width + 1300, Form1.Picture1.Height + 1300, GetDC(0), 0, 0, vbSrcCopy
Form2.Command3.Left = pt2.x + 33.34: Form2.Command3.Top = pt2.y - 20: Form2.Command4.Left = pt2.x + 66.67: Form2.Command4.Top = pt2.y - 20
dn = 1
Else:
Form1.Picture1.Width = (pt.x - pt2.x)

Form1.Picture1.Height = (pt.y - pt2.y)

BitBlt Form1.Picture1.hdc, 0 - pt2.x, 0 - pt2.y, Form1.Picture1.Width + 1300, Form1.Picture1.Height + 1300, GetDC(0), 0, 0, vbSrcCopy
Form2.Command3.Left = pt2.x + 33.34: Form2.Command3.Top = pt2.y + 20: Form2.Command4.Left = pt2.x + 66.67: Form2.Command4.Top = pt2.y + 20
dn = 1
End If

End If


Form2.Command3.Visible = True:    Form2.Command4.Visible = True
Form2.Shape1.Visible = True
End Sub
Public Sub HooK()
    hHook = SetWindowsHookEx(WH_MOUSE_LL, AddressOf MouseHook, App.hInstance, 0)
End Sub
Public Sub Save()

End Sub



Public Sub UnHooK()
  UnhookWindowsHookEx hHook
End Sub






