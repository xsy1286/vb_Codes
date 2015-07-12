Attribute VB_Name = "Module1"
Option Explicit

Public Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
  Public Declare Function CallWindowProc Lib "user32" Alias "CallWindowProcA" (ByVal lpPrevWndFunc As Long, ByVal hwnd As Long, ByVal Msg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
   Public Declare Function GetWindowLong Lib "user32" Alias "GetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long) As Long
Public Const GWL_WNDPROC = (-4)
  Public Const WM_MOVE = &H3
Public Const WM_MOUSEMOVE = &H200
Public Const WM_LBUTTONDOWN = &H201
  Public Const WM_LBUTTONUP = &H202

 Public Const WM_LBUTTONDBLCLK = &H203
  Public Const WM_RBUTTONDOWN = &H204
 Public Const WM_RBUTTONUP = &H205
   Public Const WM_RBUTTONDBLCLK = &H206
  Public Const MK_LBUTTON = &H1
  Public Const MK_RBUTTON = &H2

  Public OldProc As Long

  Public Function WndProc(ByVal hwnd As Long, ByVal Msg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
          If Msg = WM_LBUTTONDOWN Then
   Load Form2
   Unload Form2
                  Debug.Print "ÒÆ¶¯sds´°Ìå"
                
          End If
          
         WndProc = CallWindowProc(OldProc, hwnd, Msg, wParam, lParam)
  End Function

