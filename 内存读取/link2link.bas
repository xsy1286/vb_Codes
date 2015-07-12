Attribute VB_Name = "link2link"
Option Explicit
'使用到Sim_MouseKey.bas模块
Public Type Gm
X As Long
Y As Long
value As Integer
End Type
Public j1(0 To 18, 0 To 10) As Gm
Public tint As Single
Public tcik As Single
Private Function swap(ByRef a As Long, ByRef b As Long) As Boolean
Dim c As Long
If a > b Then
c = a
a = b
b = c
End If
End Function
Public Function p2plink(ByVal ax As Long, ByVal ay As Long, ByVal bx As Long, ByVal by As Long) As Boolean   '同行或同列判断，两点重合相邻判=true
Dim x1&, y1&, x2&, y2&

Dim i As Long

 p2plink = True
 
If ay = by Then

    If ax < (bx - 1) Then
    x1 = ax + 1
    x2 = bx - 1
      For i = x1 To x2
    If j1(i, ay).value > 0 Then p2plink = False: Exit Function
    Next
    
    ElseIf bx < (ax - 1) Then
    x2 = ax - 1
    x1 = bx + 1
      For i = x1 To x2
    If j1(i, ay).value > 0 Then p2plink = False: Exit Function
    Next
    
    Else
    p2plink = True: Exit Function
    
    End If
    
  

ElseIf ax = bx Then

    If ay < (by - 1) Then
     y1 = ay + 1
     y2 = by - 1
     For i = y1 To y2
    If j1(ax, i).value > 0 Then p2plink = False: Exit Function
    Next
    
    ElseIf by < (ay - 1) Then
     y2 = ay - 1
     y1 = by + 1
     For i = y1 To y2
    If j1(ax, i).value > 0 Then p2plink = False: Exit Function
    Next
    
    Else
     p2plink = True: Exit Function
    
    End If
    
   
    
Else:
 p2plink = False
End If

End Function

Private Function pnplink(ax As Long, ay As Long, bx As Long, by As Long) As Boolean
Dim ii As Long
Dim jj As Long

pnplink = False
    
    For ii = 0 To 18
    If p2plink(ax, ay, ii, ay) = True And p2plink(ii, ay, ii, by) = True And p2plink(ii, by, bx, by) = True And _
    j1(ii, ay).value = 0 And j1(ii, by).value = 0 Then
        pnplink = True: Exit Function
    ElseIf ii = ax And j1(ii, by).value = 0 And p2plink(ax, ay, ii, by) = True And p2plink(ii, by, bx, by) = True Then
        pnplink = True: Exit Function
    ElseIf ii = bx And j1(ii, ay).value = 0 And p2plink(ax, ay, ii, ay) = True And p2plink(ii, ay, ii, by) = True Then
        pnplink = True: Exit Function
    ElseIf ax = ii And bx = ii And p2plink(ax, ay, bx, by) = True Then
        pnplink = True: Exit Function
    End If
    Next ii
    
  For jj = 0 To 10
    If p2plink(ax, ay, ax, jj) = True And p2plink(ax, jj, bx, jj) = True And p2plink(bx, jj, bx, by) = True And _
    j1(ax, jj).value = 0 And j1(bx, jj).value = 0 Then
        pnplink = True: Exit Function
'    ElseIf jj = ay And j1(bx, jj).value = 0 Then '上面ii循环已包含
'        pnplink = True: Exit Function
'    ElseIf jj = by And j1(ax, jj).value = 0 Then
'        pnplink = True: Exit Function
    ElseIf ay = jj And by = jj And p2plink(ax, ay, bx, by) = True Then
        pnplink = True: Exit Function
    End If
    Next jj
    
End Function

Public Function blclink(ByVal hwd As Long) As Boolean
Dim ii As Long
Dim jj As Long
Dim k As Long: Dim n As Long
Dim x0 As Long: Dim y0 As Long

Dim Rct As RECT
Dim xS As Long: Dim yS As Long
Dim SW As Double: Dim SH As Double
Dim mx As Double: Dim my As Double

blclink = False

Call GetWindowRect(hwd, Rct)

SW = Screen.Width / Screen.TwipsPerPixelX
SH = Screen.Height / Screen.TwipsPerPixelY

 For ii = 0 To 18
  For jj = 0 To 10
    If (j1(ii, jj).value <> 0) Then
    
      k = jj * 19 + ii
      For n = (k + 1) To 208
       
       x0 = ((n + 1) Mod 19) - 1
       If x0 = -1 Then x0 = 18
       y0 = (n + 1 - x0 - 1) / 19

       
        If j1(ii, jj).value = j1(x0, y0).value Then
         If pnplink(ii, jj, x0, y0) = True Then
         Form1.Print "Find  (" & CStr(ii) & "," & CStr(jj) & ") (" & CStr(x0) & "," & CStr(y0) & ")"
         
    j1(ii, jj).value = 0: j1(x0, y0).value = 0  '*****improtant line*****
         
        xS = Rct.Left + ii * 32 + 19
        yS = Rct.Top + 180 + jj * 34 + 15

        Call SetMousePos(xS, yS)
        mouse_event MOUSEEVENTF_LEFTDOWN, 0, 0, 0, 0
        Call waittime(tcik)
        mouse_event MOUSEEVENTF_LEFTUP, 0, 0, 0, 0

       Call waittime(tint)

         xS = Rct.Left + x0 * 32 + 19
        yS = Rct.Top + 180 + y0 * 34 + 15

        Call SetMousePos(xS, yS)
        mouse_event MOUSEEVENTF_LEFTDOWN, 0, 0, 0, 0
        Call waittime(tcik)
        mouse_event MOUSEEVENTF_LEFTUP, 0, 0, 0, 0

        blclink = True
        Exit For
         End If
         
       End If
       
      Next n
     End If
     
  Next jj
 Next ii


End Function
