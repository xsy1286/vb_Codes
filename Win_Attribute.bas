Attribute VB_Name = "Win_Attribute"
Option Explicit
   Private Declare Function GetWindowLong Lib "user32" Alias "GetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long) As Long
   Private Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
   Private Declare Function SetLayeredWindowAttributes Lib "user32" (ByVal hwnd As Long, ByVal crKey As Long, ByVal bAlpha As Byte, ByVal dwFlags As Long) As Long
   Private Const WS_EX_TRANSPARENT   As Long = &H20&
   Private Const WS_EX_LAYERED = &H80000
   Private Const GWL_EXSTYLE = (-20)
   Private Const LWA_ALPHA = &H2
   Private Const LWA_COLORKEY = &H1
   
   
Public Function setAttribute(ByVal hwd As Long, ByVal clr As Long, ByVal alp As Integer, ByVal sty As Integer) As Long 'alp是不透明度，取值范围是（0,255），其中0代表全透明，255代表不透明。
                                                                                                                       'sty=0透明度+透明 =1透明 =2透明度
Dim rtn As Long
Dim s As Integer
Dim a As Integer

   If sty >= 0 And sty <= 2 Then
   rtn = GetWindowLong(hwd, GWL_EXSTYLE)
   rtn = rtn Or WS_EX_LAYERED ' Or WS_EX_TRANSPARENT
   SetWindowLong hwd, GWL_EXSTYLE, rtn
   
      If sty = 0 Then
          s = LWA_COLORKEY Or LWA_ALPHA
      ElseIf sty = 1 Then
          s = LWA_COLORKEY
      Else
          s = LWA_ALPHA
      End If
      
      If alp > 255 Then
         a = 255
      ElseIf alp < 0 Then
         a = 0
      Else
          a = alp
      End If
      
     setAttribute = SetLayeredWindowAttributes(hwd, clr, a, s)
   Else
    setAttribute = 0
   End If
End Function
