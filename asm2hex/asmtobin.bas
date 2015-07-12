Attribute VB_Name = "asmtobin"
Public Function picasm2(ByVal fst As String, ByVal sec As String, ByVal trd As String, ByVal tline As Integer) As String
pisasm2 = ""
Dim tmp As String
   tmp = ""
Select Case fst

   Case "NOP"
    tmp = "000000000000"  'decodes = 15'b000000000000000; // NOP
   Case "OPTION"
    tmp = "000000000010"  'decodes = 15'b000000100100001; // OPTION
   Case "SLEEP"
    tmp = "000000000011"  'decodes = 15'b000000000000000; // SLEEP
   Case "CLRWDT"
    tmp = "000000000100"  'decodes = 15'b000000000000000; // CLRWDT
   Case "CLRW"
    tmp = "000001000000"  'decodes = 15'b000000111010000; // CLRW
   Case "TRIS"
     If (sec = 5) Then
        tmp = "000000000101"  'decodes = 15'b000000000100010; // TRIS 5
     ElseIf (sec = 6) Then
        tmp = "000000000110"  'decodes = 15'b000000100100010; // TRIS 6
     ElseIf (sec = 7) Then
        tmp = "000000000111"  'decodes = 15'b000000100100010; // TRIS 7
    End If
    
    
   Case "MOVWF"
    tmp = "0000001" + h16tobin(sec, 5, tline) 'decodes = 15'b000000100100000; // MOVWF
   Case "CLRF"
    tmp = "0000011" + h16tobin(sec, 5, tline)  'decodes = 15'b000000110110000; // CLRF
   
   Case "SUBWF"
      If (trd = 0) Then
        tmp = "0000100" + h16tobin(sec, 5, tline) 'decodes = 15'b010010001011000; // SUBWF (d=0)
      ElseIf (trd = 1) Then
        tmp = "0000101" + h16tobin(sec, 5, tline)  'decodes = 15'b010010000111000; // SUBWF (d=1)
    End If
   Case "DECF"
    tmp = "000011" & trd & h16tobin(sec, 5, tline)    'decodes = 15'b011110001010000; // DECF  (d=0)
   Case "IORWF"
    tmp = "000100" & trd & h16tobin(sec, 5, tline)    'decodes = 15'b000100101010000; // IORWF (d=0)
   Case "ANDWF"
    tmp = "000101" & trd & h16tobin(sec, 5, tline)    'decodes = 15'b000100011010000; // ANDWF (d=0)
   Case "XORWF"
    tmp = "000110" & trd & h16tobin(sec, 5, tline)    'decodes = 15'b000100111010000; // XORWF (d=0)
   Case "ADDWF"
    tmp = "000111" & trd & h16tobin(sec, 5, tline)    'decodes = 15'b000100001011000; // ADDWF (d=0)
   Case "MOVF"
    tmp = "001000" & trd & h16tobin(sec, 5, tline)    'decodes = 15'b010100101010000; // MOVF  (d=0)
   Case "COMF"
    tmp = "001001" & trd & h16tobin(sec, 5, tline)    'decodes = 15'b010101001010000; // COMF  (d=0)
   Case "INCF"
    tmp = "001010" & trd & h16tobin(sec, 5, tline)    'decodes = 15'b011100001010000; // INCF  (d=0)
   Case "DFCFSZ"
    tmp = "001011" & trd & h16tobin(sec, 5, tline)    'decodes = 15'b011110001000000; // DECFSZ(d=0)
   Case "RRF"
    tmp = "001100" & trd & h16tobin(sec, 5, tline)    'decodes = 15'b010101011001000; // RRF   (d=0)
   Case "RLF"
    tmp = "001101" & trd & h16tobin(sec, 5, tline)    'decodes = 15'b010101101001000; // RLF   (d=0)
   Case "SWAPF"
    tmp = "001110" & trd & h16tobin(sec, 5, tline)    'decodes = 15'b010101111000000; // SWAPF (d=0)
   Case "INCFSZ"
    tmp = "001111" & trd & h16tobin(sec, 5, tline)    'decodes = 15'b011100001000000; // INCFSZ(d=0)
    
   
   
   Case "BCF"
    tmp = "0100" & h16tobin(trd, 3, tline) & h16tobin(sec, 5, tline)   'decodes = 15'b110100010100100; // BCF
   Case "BSF"
    tmp = "0101" & h16tobin(trd, 3, tline) & h16tobin(sec, 5, tline)   'decodes = 15'b110100100100000; // BSF
   Case "BTFSC"
    tmp = "0110" & h16tobin(trd, 3, tline) & h16tobin(sec, 5, tline)   'decodes = 15'b110100010000000; // BTFSC
   Case "BTFSS"
    tmp = "0111" & h16tobin(trd, 3, tline) & h16tobin(sec, 5, tline)   'decodes = 15'b110100010000000; // BTFSS
    
   Case "RETLW"
    tmp = "1000" & h16tobin(sec, 8, tline)   'decodes = 15'b101000101000000; // RETLW"
   Case "CALL"
    tmp = "1001" & h16tobin(sec, 8, tline)   'decodes = 15'b101000100000000; // CALL"
   Case "MOVLW"
    tmp = "1100" & h16tobin(sec, 8, tline)   'decodes = 15'b101000101000000; // MOVLW"
   Case "IORLW"
    tmp = "1101" & h16tobin(sec, 8, tline)   'decodes = 15'b001000101010000; // IORLW"
   Case "ANDLW"
    tmp = "1110" & h16tobin(sec, 8, tline)   'decodes = 15'b001000011010000; // ANDLW"
   Case "XORLW"
    tmp = "1111" & h16tobin(sec, 8, tline)   'decodes = 15'b001000111010000; // XORLW"
    
   Case "GOTO"
    tmp = "101" & h16tobin(sec, 9, tline)   'decodes = 15'b101000100000000; // GOTO
    
  End Select
  
    If (tmp = "") Then MsgBox "数据错误!行：" & CStr(tline): Exit Function
    Debug.Print tmp
  picasm2 = bin2h(Mid(tmp, 1, 4)) & bin2h(Mid(tmp, 5, 4)) & bin2h(Mid(tmp, 9, 4))
  Debug.Print picasm2
End Function
Public Function bin2h(two As String) As String
Dim a As Integer

'two = Mid(two, Len(two) - 3, 4)
a = Val(Mid(two, 1, 1)) * 8 + Val(Mid(two, 2, 1)) * 4 + Val(Mid(two, 3, 1)) * 2 + Val(Mid(two, 4, 1))
Debug.Print "a=  " & CStr(a)
If (a < 10) Then
  bin2h = CStr(a)
Else
 Select Case a
 Case 10
 bin2h = "A"
 Case 12
 bin2h = "B"
  Case 13
 bin2h = "C"
  Case 14
 bin2h = "D"
  Case 15
 bin2h = "E"
  Case 16
 bin2h = "F"
 End Select
End If
 Debug.Print bin2h
 
End Function


'Private Function h16tobin(inhex As String, n As Integer) As String
'用 mod 2^n - mod 2^(n-1)
'End Function

Public Function h16tobin(ByVal su As String, ByVal n As Integer, ByVal tline As Integer) As String  '16进制转换为2进制

Dim siliuTo2 As String
Dim k As String
Dim m As Integer

If ((n Mod 4) <> 0) Then
m = ((4 - (n Mod 4)) + n)
Else
m = n
End If

If (Len(su) < (m / 4)) Then

For i = 1 To ((m / 4) - Len(su))
su = "0" & su
Next i
Debug.Print "su=  " & su

End If

For i = 1 To Len(su)
k = Mid(su, i, 1)
Select Case k
Case "0"
siliuTo2 = siliuTo2 & "0000"
Case "1"
siliuTo2 = siliuTo2 & "0001"
Case "2"
siliuTo2 = siliuTo2 & "0010"
Case "3"
siliuTo2 = siliuTo2 & "0011"
Case "4"
siliuTo2 = siliuTo2 & "0100"
Case "5"
siliuTo2 = siliuTo2 & "0101"
Case "6"
siliuTo2 = siliuTo2 & "0110"
Case "7"
siliuTo2 = siliuTo2 & "0111"
Case "8"
siliuTo2 = siliuTo2 & "1000"
Case "9"
siliuTo2 = siliuTo2 & "1001"
Case "A"
siliuTo2 = siliuTo2 & "1010"
Case "B"
siliuTo2 = siliuTo2 & "1011"
Case "C"
siliuTo2 = siliuTo2 & "1100"
Case "D"
siliuTo2 = siliuTo2 & "1101"
Case "E"
siliuTo2 = siliuTo2 & "1110"
Case "F"
siliuTo2 = siliuTo2 & "1111"

Case Else
MsgBox "数据错误!行：" & CStr(tline)
Exit Function

End Select
Next

h16tobin = Mid(siliuTo2, Len(siliuTo2) - n + 1, n)

End Function
