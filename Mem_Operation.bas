Attribute VB_Name = "MemOperation"
Option Explicit
Dim dat() As Byte
Private Declare Function FindWindow Lib "user32" Alias "FindWindowA" (ByVal lpClassName As String, ByVal lpWindowName As String) As Long
Public Declare Function GetWindowText Lib "user32" Alias "GetWindowTextA" (ByVal hwnd As Long, ByVal lpString As String, ByVal cch As Long) As Long
Public Declare Function GetNextWindow Lib "user32" Alias "GetWindow" (ByVal hwnd As Long, ByVal wFlag As Long) As Long
Public Declare Function GetForegroundWindow Lib "user32" () As Long

'Public Declare Function OpenProcess Lib "kernel32" (ByVal dwDesiredAccess As Long, ByVal bInheritHandle As Long, ByVal dwProcessId As Long) As Long
Private Declare Function OpenProcess Lib "kernel32" (ByVal dwDesiredAccess As Long, ByVal bInheritHandle As Long, ByVal dwProcessId As Long) As Long
Private Declare Function CloseHandle Lib "kernel32" (ByVal hObject As Long) As Long

Private Declare Function GetWindowThreadProcessId Lib "user32" (ByVal hwnd As Long, lpdwProcessId As Long) As Long
                                                                                                                 ' Read必为 ByVal lpBuffer
Private Declare Function ReadProcessMemory Lib "kernel32" (ByVal hProcess As Long, ByVal lpBaseAddress As Long, ByVal lpBuffer As Any, ByVal nSize As Long, ByRef lpNumberOfBytesWritten As Long) As Long
Private Declare Function WriteProcessMemory Lib "kernel32" (ByVal hProcess As Long, ByVal lpBaseAddress As Any, ByVal lpBuffer As Any, ByVal nSize As Long, lpNumberOfBytesWritten As Long) As Long

 Const PROCESS_VM_OPERATION = &H8&
 Const PROCESS_VM_READ = &H10&
 Const PROCESS_VM_WRITE = &H20&
 Const PROCESS_ALL_ACCESS = &H1F0FFF

Private Declare Function GetWindowRect Lib "user32" (ByVal hwnd As Long, lpRect As RECT) As Long  '，Right跟底部都需减1才是真实值

Public Sub P2M(nam As String, ByVal lgh As Long, ByVal address As Long)
Dim PID&
Dim btData&
Dim h As Long
Dim ph As Long

h = FindWindow(vbNullString, nam)

If h = 0 Then Debug.Print "hwnd=0": Exit Sub

GetWindowThreadProcessId h, PID

ph = OpenProcess(PROCESS_ALL_ACCESS, False, PID)


Debug.Print "subopenPro succeed" & " PID:" & CStr(PID) & " handle:" & CStr(ph)

'Dim k(0 To 208) As Byte
'Dim n%
'For n = 0 To 206
'k(n) = 0
'Next n
'k(207) = &H15
'k(208) = &H15
'
ReDim dat(lgh) As Byte
ReadProcessMemory ph, address, VarPtr(dat(0)), lgh, 0&

Dim i&
Dim j&
For i = 0 To 10
    For j = 0 To 18
        If (j <> 18) Then
         Form1.Print dat(i * 19 + j);
        Else
         Form1.Print dat(i * 19 + j)
        End If
    Next
Next


End Sub
Public Function myProcessopen(nam As String, ByRef ph As Long) As Long
Dim PID&
Dim btData&
Dim h As Long
'Dim ph As Long

h = FindWindow(vbNullString, nam)

If h = 0 Then Debug.Print "hwnd=0": Exit Function

GetWindowThreadProcessId h, PID

ph = OpenProcess(PROCESS_ALL_ACCESS, False, PID)


myProcessopen = h

Debug.Print "subopenPro succeed" & " PID:" & CStr(PID) & " handle:" & CStr(ph)


End Function

Public Function myPrsOpenByhWnd(h As Long, ByRef ph As Long) As Long
Dim PID&
Dim btData&

'Dim ph As Long


If h = 0 Then Debug.Print "hwnd=0": MsgBox "hWnd 无效": Exit Function

GetWindowThreadProcessId h, PID

ph = OpenProcess(PROCESS_ALL_ACCESS, False, PID)


myPrsOpenByhWnd = h

Debug.Print "subopenPro succeed" & " PID:" & CStr(PID) & " handle:" & CStr(ph)


End Function

Public Function getMem(ByVal address As Long, ByVal lgh As Long, ByVal phl As Long) As Byte()
On Error GoTo errlog
ReDim dat(lgh) As Byte
Dim r As Long
r = ReadProcessMemory(phl, address, VarPtr(dat(0)), lgh, 0&)

If r = 0 Then
    Debug.Print "read Memory fail" & "-" & "  LOG address:" & CStr(address) & "length:" & CStr(lgh)
Else
    Debug.Print "read Memory succeed" & "-" & " LOG address:" & CStr(address) & "length:" & CStr(lgh)
End If
'WriteProcessMemory ph, &H129F78, VarPtr(k(0)), 209, 0&
getMem = dat


Exit Function
errlog:
Call whenErr(Err.Number, "MemRead", "Read")
End Function

Public Function wrtMem(ByVal address As Long, ByVal lgh As Long, ByVal phandle As Long, ByRef con() As Byte) As Long '注意用时,con()数组要比lgh长
'内存要写必须得改 OpenProcess 中的权限参数
On Error GoTo errlog

wrtMem = WriteProcessMemory(phandle, address, VarPtr(con(LBound(con))), lgh, 0&)
'WriteProcessMemory ph, &H129F78, VarPtr(k(0)), 209, 0&

Debug.Print "write succeed"

Exit Function
errlog:
Call whenErr(Err.Number, "MemRead", "Write")
End Function

Public Function forhandleclose(ByVal ph As Long) '在OpenProcess（API）之后必须CloseHandle（API），也就是在自定义的myProcessopen传出的handle用完之后得，用此forhandleclose函数回收
CloseHandle ph
End Function

Public Function nameToHwndEx(ByVal nam As String, ci() As Long, ByVal depth As Integer) As Long '大于等于的窗口标题栏文本也会被匹配
                                                                'depth建议设为100以上
Dim h&
Dim str As String
  str = Space(255) '应该此句必要
Dim i As Integer
Dim l As Integer
    i = 0
h = GetForegroundWindow
Do While (i <= depth) And (h <> 0)
   l = GetWindowText(h, str, 256)
    If Mid(str, 1, Len(nam)) = nam Then ci(i) = h: i = i + 1
   h = GetNextWindow(h, 2) '仅知2为向下查找句柄
Loop
nameToHwndEx = i
End Function

Public Function WndWid(ByVal fhwd As Long) As RECT
Dim Rct As RECT
Call GetWindowRect(fhwd, Rct)
Debug.Print CStr(fhwd)
Debug.Print "left:  " & CStr(Rct.Left) & "  " & "top:  " & CStr(Rct.Top)
Debug.Print "right:  " & CStr(Rct.Right) & "  " & "bottom:  " & CStr(Rct.Bottom)
WndWid = Rct

End Function

