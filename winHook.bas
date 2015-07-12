Attribute VB_Name = "winHook"
Option Explicit

Public Const WH_KEYBOARD_LL = 13
Public Const WH_MOUSE_LL = 14
Public Declare Function SetWindowsHookEx Lib "user32" Alias "SetWindowsHookExA" (ByVal idHook As Long, ByVal lpfn As Long, ByVal hmod As Long, ByVal dwThreadId As Long) As Long
Public Declare Function UnhookWindowsHookEx Lib "user32" (ByVal hHook As Long) As Long
Public Declare Function CallNextHookEx Lib "user32" (ByVal hHook As Long, ByVal nCode As Long, ByVal wParam As Long, lParam As Any) As Long

Public Declare Sub CopyMemory Lib "kernel32" _
                          Alias "RtlMoveMemory" _
                          (Destination As Any, _
                          Source As Any, _
                          ByVal Length As Long)
Private Type POINTAPI
    x As Long
    y As Long
End Type

Private Type MSLLHOOKSTRUCT  'module中API不能private，但Type可以
    pt As POINTAPI
    mouseData As Long
    Flags As Long
    time As Long
    dwExtraInfo As Long
End Type
Private hHook As Long
Public Sub hook()

    hHook = SetWindowsHookEx(WH_MOUSE_LL, AddressOf MouseHook, App.hInstance, 0)  '第一个参数是不是把Hook只限制在鼠标

End Sub

Public Sub UnHooK()

    If hHook <> 0 Then
        UnhookWindowsHookEx hHook
    End If

End Sub

Public Function MouseHook(ByVal nCode As Long, _
                          ByVal wParam As Long, _
                          ByVal lParam As Long) As Long

    Dim mhs     As MSLLHOOKSTRUCT

    Static m_pt As POINTAPI

    Call CopyMemory(mhs, ByVal lParam, LenB(mhs))
    m_pt = mhs.pt
        
    '    If wParam = WM_LBUTTONDOWN Then
    '
    '    End If
'******************************************************
'example for call your own Hook Action
    'Call Form1.fnHook(m_pt.x, m_pt.y, wParam)
'******************************************************

    MouseHook = CallNextHookEx(hHook, nCode, wParam, lParam)
End Function

