Attribute VB_Name = "Module1"
   'module ��ע��Form1�������ҪForm dot
Option Explicit
#Const d_hook = 0
Public Declare Function PrintWindow Lib "user32" (ByVal hwnd As Long, ByVal hdcBlt As Long, ByVal nFlags As Long) As Long
Public Declare Function BitBlt Lib "gdi32" (ByVal hDestDC As Long, ByVal X As Long, ByVal Y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal dwRop As Long) As Long
Public Declare Function StretchBlt Lib "gdi32" (ByVal hdc As Long, ByVal X As Long, ByVal Y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal nSrcWidth As Long, ByVal nSrcHeight As Long, ByVal dwRop As Long) As Long
Public Declare Function GetDC Lib "user32" (ByVal hwnd As Long) As Long
Public Declare Function GetCursorPos Lib "user32" (lpPonit As POINTAPI) As Long

Public Declare Function SetWindowsHookEx Lib "user32" Alias "SetWindowsHookExA" (ByVal idHook As Long, ByVal lpfn As Long, ByVal hmod As Long, ByVal dwThreadId As Long) As Long
Public Declare Function UnhookWindowsHookEx Lib "user32" (ByVal hHook As Long) As Long
Public Declare Function CallNextHookEx Lib "user32" (ByVal hHook As Long, ByVal nCode As Long, ByVal wParam As Long, lParam As Any) As Long
Public Declare Sub CopyMemory Lib "kernel32" _
                          Alias "RtlMoveMemory" _
                          (Destination As Any, _
                          Source As Any, _
                          ByVal Length As Long)
Public l, h As Long
Public tt As Integer
Public bb As Boolean
Public d, dd As Integer
Public Type POINTAPI
    X As Long
    Y As Long
End Type
Private ret As Long

Private Type MSLLHOOKSTRUCT  'module��API����private����Type����
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
'��Ϣ
Public Const HC_ACTION = 0
Public Const HC_SYSMODALOFF = 5
Public Const HC_SYSMODALON = 4
'�����Ϣ
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
Public ptf4 As POINTAPI
Public lHook As Long
Public hHook As Long
Public shapeexist As Boolean
'��hook��ʹ��
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

      '  Debug.Print "�������    ����:" & pt.x & "  "; pt.y
    End If
    
    
    
     If wParam = WM_LBUTTONUP And d = 1 Then
          Call CopyMemory(mhs, ByVal lParam, LenB(mhs))
        pt2 = mhs.pt
     d = 0
    
        'Debug.Print "�����    ����:" & pt2.x & "  "; pt2.y



            
      
           Call ShapeDraw
         If pt.X <> pt2.X Or pt2.Y <> pt.Y Then
         '��false ��true��ֹ��ʾϵͳ�Դ���
             tt = 1
           Call Draw
         ' End If
          End If
    End If


If wParam = WM_MOUSEMOVE And d = 1 Then
        Call CopyMemory(mhs, ByVal lParam, LenB(mhs))     '����ʹ��Len ���������ַ��������ַ����� ����Dim n As Integer ����n=LenB("Hello world") //����11
        If dn = 1 Then dn = 0
        pt2 = mhs.pt
       
           Call ShapeDraw
    End If

    Call CallNextHookEx(hHook, nCode, wParam, lParam)
End Function


Public Sub ShapeDraw()

If pt2.X > pt.X Then
      If pt2.Y > pt.Y Then
      Form2.Shape1.Left = pt.X - 2: Form2.Shape1.Top = pt.Y - 2
      'Form2.Command1.Caption = Str(pt.X): Form2.Command1.Caption = Str(pt.Y)
Form2.Shape1.Width = (pt2.X - pt.X) + 2

Form2.Shape1.Height = (pt2.Y - pt.Y) + 2

Else:
Form2.Shape1.Left = pt.X - 2: Form2.Shape1.Top = pt2.Y - 2
Form2.Shape1.Width = (pt2.X - pt.X) + 2
Form2.Shape1.Height = (pt.Y - pt2.Y) + 2

End If

End If


      If pt2.X < pt.X Then
      If pt2.Y > pt.Y Then
      Form2.Shape1.Left = pt2.X - 2: Form2.Shape1.Top = pt.Y - 2
Form2.Shape1.Width = (pt.X - pt2.X) + 2
Form2.Shape1.Height = (pt2.Y - pt.Y) + 2

Else:

Form2.Shape1.Width = (pt.X - pt2.X) + 2
Form2.Shape1.Height = (pt.Y - pt2.Y) + 2
Form2.Shape1.Left = pt2.X - 2: Form2.Shape1.Top = pt2.Y - 2
End If

End If
shapeexist = True
End Sub

Public Sub Draw()
Form1.Picture1.Cls 'ע��������ػ�
     Form2.Command3.Visible = False
      Form2.Command4.Visible = False


Call delay(150)
      
      'module ��ע��Form1�������ҪForm dot
      If pt2.X > pt.X Then
      
       If pt2.Y > pt.Y Then

Form1.Picture1.Width = (pt2.X - pt.X)

Form1.Picture1.Height = (pt2.Y - pt.Y)
l = (pt2.X - pt.X)
h = (pt2.Y - pt.Y)
Form2.Label2.Caption = Str(l) + "   " + Str(h)
                            '��Ϊ���ص����
                               '��    ��
                            
 BitBlt Form1.Picture1.hdc, (0 - pt.X), (0 - pt.Y), Form1.Picture1.Width + 1300, Form1.Picture1.Height + 1300, GetDC(0), 0, 0, vbSrcCopy
'Debug.Print "r=" & Str(0 - pt.X) & Str(0 - pt.Y)
'Form1.Picture1.PaintPicture Form2.Picture, 0, 0, Form1.Picture1.Width, Form1.Picture1.Height, pt.X  , pt.Y  , Screen.Width, Screen.Height, vbSrcCopy
  Form2.Command3.Left = pt2.X - 66.67: Form2.Command3.Top = pt2.Y - 20: Form2.Command4.Left = pt2.X - 33.34: Form2.Command4.Top = pt2.Y - 20
dn = 1
Else:
Form1.Picture1.Width = (pt2.X - pt.X)
Form1.Picture1.Height = (pt.Y - pt2.Y)

BitBlt Form1.Picture1.hdc, 0 - pt.X, 0 - pt2.Y, Form1.Picture1.Width + 1300, Form1.Picture1.Height + 1300, GetDC(0), 0, 0, vbSrcCopy
Form2.Command3.Left = pt2.X - 66.67: Form2.Command3.Top = pt2.Y + 20: Form2.Command4.Left = pt2.X - 33.34: Form2.Command4.Top = pt2.Y + 20
dn = 1
End If

End If


      If pt2.X < pt.X Then
      If pt2.Y > pt.Y Then
Form1.Picture1.Width = (pt.X - pt2.X)
Form1.Picture1.Height = (pt2.Y - pt.Y)

BitBlt Form1.Picture1.hdc, 0 - pt2.X, 0 - pt.Y, Form1.Picture1.Width + 1300, Form1.Picture1.Height + 1300, GetDC(0), 0, 0, vbSrcCopy
Form2.Command3.Left = pt2.X + 33.34: Form2.Command3.Top = pt2.Y - 20: Form2.Command4.Left = pt2.X + 66.67: Form2.Command4.Top = pt2.Y - 20
dn = 1
Else:
Form1.Picture1.Width = (pt.X - pt2.X)

Form1.Picture1.Height = (pt.Y - pt2.Y)

BitBlt Form1.Picture1.hdc, 0 - pt2.X, 0 - pt2.Y, Form1.Picture1.Width + 1300, Form1.Picture1.Height + 1300, GetDC(0), 0, 0, vbSrcCopy
Form2.Command3.Left = pt2.X + 33.34: Form2.Command3.Top = pt2.Y + 20: Form2.Command4.Left = pt2.X + 66.67: Form2.Command4.Top = pt2.Y + 20
dn = 1
End If

End If


Form2.Command3.Visible = True:    Form2.Command4.Visible = True
Form2.Shape1.Visible = True
End Sub

Public Sub hook()
#If d_hook Then
hHook = SetWindowsHookEx(WH_MOUSE_LL, AddressOf MouseHook, App.hInstance, 0)
#End If
End Sub
Public Sub UnHooK()
#If d_hook Then
UnhookWindowsHookEx hHook
#End If
End Sub

Public Function MouseHookf4(ByVal nCode As Long, _
                       ByVal wParam As Long, _
                       ByVal lParam As Long) As Long

    Dim mhs As MSLLHOOKSTRUCT

 If wParam <> WM_MOUSEMOVE And wParam <> WM_RBUTTONUP Then

If wParam = WM_LBUTTONDOWN Or wParam = WM_LBUTTONUP Then
   Call CopyMemory(mhs, ByVal lParam, LenB(mhs))
        ptf4 = mhs.pt
        
If _
 ptf4.X > Form4.Left / Screen.TwipsPerPixelX And _
  ptf4.X < (Form4.Left + Form4.Width) / Screen.TwipsPerPixelX And _
  ptf4.Y > Form4.Top / Screen.TwipsPerPixelY And _
  ptf4.Y < (Form4.Top + Form4.Height) / Screen.TwipsPerPixelY _
Then

Else:
 Form4.Hide
      Unload Form4

End If


Else:
        Form4.Hide
      Unload Form4
      End If
'function name=1 '�Ե������Ϣ
'function name = CallNextHookEx(�¸�hook,��hook, wParam, ByVal lParam) '������һ������
End If
End Sub
