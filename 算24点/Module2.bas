Attribute VB_Name = "For_24"
Public Sub col(ByVal k As Integer, ByVal x1 As Integer, ByVal x2 As Integer, ByVal x3 As Integer, ByVal x4 As Integer, ByRef t As Integer, ByRef s() As String) '���鲻�ñ�������
'On Error Resume Next  'For c/(a+b)  a+b==0

Dim a, b, c, d As Integer


If k = 1 Then a = x1: b = x2: c = x3: d = x4
If k = 2 Then a = x1: b = x2: c = x4: d = x3
If k = 3 Then a = x1: b = x3: c = x2: d = x4
If k = 4 Then a = x1: b = x3: c = x4: d = x2
If k = 5 Then a = x1: b = x4: c = x2: d = x3
If k = 6 Then a = x1: b = x4: c = x3: d = x2

If k = 7 Then a = x2: b = x1: c = x3: d = x4
If k = 9 Then a = x2: b = x1: c = x4: d = x3
If k = 10 Then a = x2: b = x3: c = x1: d = x4
If k = 11 Then a = x2: b = x3: c = x4: d = x1
If k = 12 Then a = x2: b = x4: c = x1: d = x3
If k = 13 Then a = x2: b = x4: c = x3: d = x1

If k = 14 Then a = x3: b = x1: c = x4: d = x2
If k = 15 Then a = x3: b = x1: c = x2: d = x4
If k = 16 Then a = x3: b = x2: c = x1: d = x4
If k = 17 Then a = x3: b = x2: c = x4: d = x1
If k = 18 Then a = x3: b = x4: c = x1: d = x2
If k = 19 Then a = x3: b = x4: c = x2: d = x1

If k = 20 Then a = x4: b = x1: c = x3: d = x2
If k = 21 Then a = x4: b = x1: c = x2: d = x3
If k = 22 Then a = x4: b = x2: c = x1: d = x3
If k = 23 Then a = x4: b = x2: c = x3: d = x1
If k = 24 Then a = x4: b = x3: c = x1: d = x2
If k = 8 Then a = x4: b = x3: c = x2: d = x1

'�������ŵļ�

If a + b + c + d = 24 Then t = t + 1: s(t) = CStr(a) + "+" + CStr(b) + "+" + CStr(c) + "+" + CStr(d) + "=" + "24"
If a + b + c - d = 24 Then t = t + 1: s(t) = CStr(a) + "+" + CStr(b) + "+" + CStr(c) + "-" + CStr(d) + "=" + "24"
If a + b - c - d = 24 Then t = t + 1: s(t) = CStr(a) + "+" + CStr(b) + "-" + CStr(c) + "-" + CStr(d) + "=" + "24"

If a + b + c * d = 24 Then t = t + 1: s(t) = CStr(a) + "+" + CStr(b) + "+" + CStr(c) + "*" + CStr(d) + "=" + "24"
If a + b + c / d = 24 Then t = t + 1: s(t) = CStr(a) + "+" + CStr(b) + "+" + CStr(c) + "/" + CStr(d) + "=" + "24"
If a + b - c * d = 24 Then t = t + 1: s(t) = CStr(a) + "+" + CStr(b) + "-" + CStr(c) + "*" + CStr(d) + "=" + "24"
If a + b - c / d = 24 Then t = t + 1: s(t) = CStr(a) + "+" + CStr(b) + "-" + CStr(c) + "/" + CStr(d) + "=" + "24"
If a + b * c * d = 24 Then t = t + 1: s(t) = CStr(a) + "+" + CStr(b) + "*" + CStr(c) + "*" + CStr(d) + "=" + "24"
If a + b / c * d = 24 Then t = t + 1: s(t) = CStr(a) + "+" + CStr(b) + "/" + CStr(c) + "*" + CStr(d) + "=" + "24"
If a + b / c / d = 24 Then t = t + 1: s(t) = CStr(a) + "+" + CStr(b) + "/" + CStr(c) + "/" + CStr(d) + "=" + "24"

If a + b * c - d = 24 Then t = t + 1: s(t) = CStr(a) + "+" + CStr(b) + "*" + CStr(c) + "-" + CStr(d) + "=" + "24"
If a + b / c - d = 24 Then t = t + 1: s(t) = CStr(a) + "+" + CStr(b) + "/" + CStr(c) + "-" + CStr(d) + "=" + "24"

'�������ŵļ�
If a - b + c * d = 24 Then t = t + 1: s(t) = CStr(a) + "-" + CStr(b) + "+" + CStr(c) + "*" + CStr(d) + "=" + "24"
If a - b + c / d = 24 Then t = t + 1: s(t) = CStr(a) + "-" + CStr(b) + "+" + CStr(c) + "/" + CStr(d) + "=" + "24"
If a - b - c * d = 24 Then t = t + 1: s(t) = CStr(a) + "-" + CStr(b) + "-" + CStr(c) + "*" + CStr(d) + "=" + "24"
If a - b - c / d = 24 Then t = t + 1: s(t) = CStr(a) + "-" + CStr(b) + "-" + CStr(c) + "/" + CStr(d) + "=" + "24"
'If a - b * c - d = 24 Then t = t + 1 :s(t) = CStr(a) + "-" + CStr(b) + "*" + CStr(c) + "-" + CStr(d) + "=" + "24"
If a - b * c * d = 24 Then t = t + 1: s(t) = CStr(a) + "-" + CStr(b) + "*" + CStr(c) + "*" + CStr(d) + "=" + "24"
If a - b * c * d = 24 Then t = t + 1: s(t) = CStr(a) + "-" + CStr(b) + "*" + CStr(c) + "*" + CStr(d) + "=" + "24"
If a - b * c / d = 24 Then t = t + 1: s(t) = CStr(a) + "-" + CStr(b) + "*" + CStr(c) + "/" + CStr(d) + "=" + "24"
If a - b / c - d = 24 Then t = t + 1: s(t) = CStr(a) + "-" + CStr(b) + "/" + CStr(c) + "-" + CStr(d) + "=" + "24"
If a - b / c / d = 24 Then t = t + 1: s(t) = CStr(a) + "-" + CStr(b) + "/" + CStr(c) + "/" + CStr(d) + "=" + "24"
 
 '�������ŵĳˣ��ɸ��Ƽӵ�   '���ҷֳ˻��
If a * b + c * d = 24 Then t = t + 1: s(t) = CStr(a) + "*" + CStr(b) + "+" + CStr(c) + "*" + CStr(d) + "=" + "24"
If a * b + c / d = 24 Then t = t + 1: s(t) = CStr(a) + "*" + CStr(b) + "+" + CStr(c) + "/" + CStr(d) + "=" + "24"
If a * b - c * d = 24 Then t = t + 1: s(t) = CStr(a) + "*" + CStr(b) + "-" + CStr(c) + "*" + CStr(d) + "=" + "24"
If a * b - c / d = 24 Then t = t + 1: s(t) = CStr(a) + "*" + CStr(b) + "-" + CStr(c) + "/" + CStr(d) + "=" + "24"
If a * b - c - d = 24 Then t = t + 1: s(t) = CStr(a) + "*" + CStr(b) + "-" + CStr(c) + "-" + CStr(d) + "=" + "24"

If a * b * c + d = 24 Then t = t + 1: s(t) = CStr(a) + "*" + CStr(b) + "*" + CStr(c) + "+" + CStr(d) + "=" + "24"
If a * b * c - d = 24 Then t = t + 1: s(t) = CStr(a) + "*" + CStr(b) + "*" + CStr(c) + "-" + CStr(d) + "=" + "24"
If a * b / c + d = 24 Then t = t + 1: s(t) = CStr(a) + "*" + CStr(b) + "/" + CStr(c) + "+" + CStr(d) + "=" + "24"
If a * b / c - d = 24 Then t = t + 1: s(t) = CStr(a) + "*" + CStr(b) + "/" + CStr(c) + "-" + CStr(d) + "=" + "24"

'ȫ�˳�  '��Χ����
If a * b * c * d = 24 Then t = t + 1: s(t) = CStr(a) + "*" + CStr(b) + "*" + CStr(c) + "*" + CStr(d) + "=" + "24"
If a * b * c / d = 24 Then t = t + 1: s(t) = CStr(a) + "*" + CStr(b) + "*" + CStr(c) + "/" + CStr(d) + "=" + "24"
If a * b / c / d = 24 Then t = t + 1: s(t) = CStr(a) + "*" + CStr(b) + "/" + CStr(c) + "/" + CStr(d) + "=" + "24"
If a / b / c / d = 24 Then t = t + 1: s(t) = CStr(a) + "/" + CStr(b) + "/" + CStr(c) + "/" + CStr(d) + "=" + "24"

 '�������ŵĳ�
If a / b + c * d = 24 Then t = t + 1: s(t) = CStr(a) + "/" + CStr(b) + "+" + CStr(c) + "*" + CStr(d) + "=" + "24"
If a / b + c / d = 24 Then t = t + 1: s(t) = CStr(a) + "/" + CStr(b) + "+" + CStr(c) + "/" + CStr(d) + "=" + "24"
If a / b - c * d = 24 Then t = t + 1: s(t) = CStr(a) + "/" + CStr(b) + "-" + CStr(c) + "*" + CStr(d) + "=" + "24"
If a / b - c / d = 24 Then t = t + 1: s(t) = CStr(a) + "/" + CStr(b) + "-" + CStr(c) + "/" + CStr(d) + "=" + "24"
If a / b - c - d = 24 Then t = t + 1: s(t) = CStr(a) + "/" + CStr(b) + "-" + CStr(c) + "-" + CStr(d) + "=" + "24"

If a / b / c + d = 24 Then t = t + 1: s(t) = CStr(a) + "/" + CStr(b) + "/" + CStr(c) + "+" + CStr(d) + "=" + "24"
If a / b / c - d = 24 Then t = t + 1: s(t) = CStr(a) + "/" + CStr(b) + "/" + CStr(c) + "-" + CStr(d) + "=" + "24"

'�����ŵģ��Ӽ��˳������ɸ������治�����ŵģ��Ӽ��˳���
If (a + b) * c * d = 24 Then t = t + 1: s(t) = "(" + CStr(a) + "+" + CStr(b) + ")" + "*" + CStr(c) + "*" + CStr(d) + "=" + "24"
If (a + b) * c / d = 24 Then t = t + 1: s(t) = "(" + CStr(a) + "+" + CStr(b) + ")" + "*" + CStr(c) + "/" + CStr(d) + "=" + "24"
If (a + b) / c * d = 24 Then t = t + 1: s(t) = "(" + CStr(a) + "+" + CStr(b) + ")" + "/" + CStr(c) + "*" + CStr(d) + "=" + "24"
If (a + b) / c / d = 24 Then t = t + 1: s(t) = "(" + CStr(a) + "+" + CStr(b) + ")" + "/" + CStr(c) + "/" + CStr(d) + "=" + "24"
If (a - b) * c * d = 24 Then t = t + 1: s(t) = "(" + CStr(a) + "-" + CStr(b) + ")" + "*" + CStr(c) + "*" + CStr(d) + "=" + "24"
If (a - b) * c / d = 24 Then t = t + 1: s(t) = "(" + CStr(a) + "-" + CStr(b) + ")" + "*" + CStr(c) + "/" + CStr(d) + "=" + "24"
If (a - b) / c * d = 24 Then t = t + 1: s(t) = "(" + CStr(a) + "-" + CStr(b) + ")" + "/" + CStr(c) + "*" + CStr(d) + "=" + "24"
If (a - b) / c / d = 24 Then t = t + 1: s(t) = "(" + CStr(a) + "-" + CStr(b) + ")" + "/" + CStr(c) + "/" + CStr(d) + "=" + "24"
'��������������ܴ�  so,����������������ܴ�
If (a + b) * c - d = 24 Then t = t + 1: s(t) = "(" + CStr(a) + "+" + CStr(b) + ")" + "*" + CStr(c) + "-" + CStr(d) + "=" + "24"
If (a + b) / c - d = 24 Then t = t + 1: s(t) = "(" + CStr(a) + "+" + CStr(b) + ")" + "/" + CStr(c) + "-" + CStr(d) + "=" + "24"

If (a + b) / c * d = 24 Then t = t + 1: s(t) = "(" + CStr(a) + "+" + CStr(b) + ")" + "/" + CStr(c) + "*" + CStr(d) + "=" + "24"
If (a + b) / c / d = 24 Then t = t + 1: s(t) = "(" + CStr(a) + "+" + CStr(b) + ")" + "/" + CStr(c) + "/" + CStr(d) + "=" + "24"

If (a - b) * c - d = 24 Then t = t + 1: s(t) = "(" + CStr(a) + "-" + CStr(b) + ")" + "*" + CStr(c) + "-" + CStr(d) + "=" + "24"
If (a - b) / c - d = 24 Then t = t + 1: s(t) = "(" + CStr(a) + "-" + CStr(b) + ")" + "/" + CStr(c) + "-" + CStr(d) + "=" + "24"
If (c - d) <> 0 Then
If a - b / (c - d) = 24 Then t = t + 1: s(t) = CStr(a) + "-" + CStr(b) + "/" + "(" + CStr(c) + "-" + CStr(d) + ")" + "=" + "24"
End If
If (c + d) <> 0 Then
If a - b / (c + d) = 24 Then t = t + 1: s(t) = CStr(a) + "-" + CStr(b) + "/" + "(" + CStr(c) + "+" + CStr(d) + ")" + "=" + "24"
End If
If (a - b) / c * d = 24 Then t = t + 1: s(t) = "(" + CStr(a) + "-" + CStr(b) + ")" + "/" + CStr(c) + "*" + CStr(d) + "=" + "24"
If (a - b) / c / d = 24 Then t = t + 1: s(t) = "(" + CStr(a) + "-" + CStr(b) + ")" + "/" + CStr(c) + "/" + CStr(d) + "=" + "24"


If a + (b + c) * d = 24 Then t = t + 1: s(t) = CStr(a) + "+" + "(" + CStr(b) + "+" + CStr(c) + ")" + "*" + CStr(d) + "=" + "24"
If a + (b + c) / d = 24 Then t = t + 1: s(t) = CStr(a) + "+" + "(" + CStr(b) + "+" + CStr(c) + ")" + "/" + CStr(d) + "=" + "24"
If a - (b + c) * d = 24 Then t = t + 1: s(t) = CStr(a) + "-" + "(" + CStr(b) + "+" + CStr(c) + ")" + "*" + CStr(d) + "=" + "24"
If a - (b + c) / d = 24 Then t = t + 1: s(t) = CStr(a) + "-" + "(" + CStr(b) + "+" + CStr(c) + ")" + "/" + CStr(d) + "=" + "24"
If a * (b + c) * d = 24 Then t = t + 1: s(t) = CStr(a) + "*" + "(" + CStr(b) + "+" + CStr(c) + ")" + "*" + CStr(d) + "=" + "24"
If a * (b + c) / d = 24 Then t = t + 1: s(t) = CStr(a) + "*" + "(" + CStr(b) + "+" + CStr(c) + ")" + "/" + CStr(d) + "=" + "24"
If a / (b + c) + d = 24 Then t = t + 1: s(t) = CStr(a) + "/" + "(" + CStr(b) + "+" + CStr(c) + ")" + "+" + CStr(d) + "=" + "24"
If a / (b + c) - d = 24 Then t = t + 1: s(t) = CStr(a) + "/" + "(" + CStr(b) + "+" + CStr(c) + ")" + "-" + CStr(d) + "=" + "24"
If a / (b + c) * d = 24 Then t = t + 1: s(t) = CStr(a) + "/" + "(" + CStr(b) + "+" + CStr(c) + ")" + "*" + CStr(d) + "=" + "24"
If a / (b + c) / d = 24 Then t = t + 1: s(t) = CStr(a) + "/" + "(" + CStr(b) + "+" + CStr(c) + ")" + "/" + CStr(d) + "=" + "24"

If a + (b - c) * d = 24 Then t = t + 1: s(t) = CStr(a) + "+" + "(" + CStr(b) + "-" + CStr(c) + ")" + "*" + CStr(d) + "=" + "24"
If a + (b - c) / d = 24 Then t = t + 1: s(t) = CStr(a) + "+" + "(" + CStr(b) + "-" + CStr(c) + ")" + "/" + CStr(d) + "=" + "24"
If a - (b - c) * d = 24 Then t = t + 1: s(t) = CStr(a) + "-" + "(" + CStr(b) + "-" + CStr(c) + ")" + "*" + CStr(d) + "=" + "24"
If a - (b - c) / d = 24 Then t = t + 1: s(t) = CStr(a) + "-" + "(" + CStr(b) + "-" + CStr(c) + ")" + "/" + CStr(d) + "=" + "24"
If a * (b - c) * d = 24 Then t = t + 1: s(t) = CStr(a) + "*" + "(" + CStr(b) + "-" + CStr(c) + ")" + "*" + CStr(d) + "=" + "24"
If a * (b - c) / d = 24 Then t = t + 1: s(t) = CStr(a) + "*" + "(" + CStr(b) + "-" + CStr(c) + ")" + "/" + CStr(d) + "=" + "24"
If b <> c Then
If a / (b - c) + d = 24 Then t = t + 1: s(t) = CStr(a) + "/" + "(" + CStr(b) + "-" + CStr(c) + ")" + "+" + CStr(d) + "=" + "24"
If a / (b - c) - d = 24 Then t = t + 1: s(t) = CStr(a) + "/" + "(" + CStr(b) + "-" + CStr(c) + ")" + "-" + CStr(d) + "=" + "24"
If a / (b - c) * d = 24 Then t = t + 1: s(t) = CStr(a) + "/" + "(" + CStr(b) + "-" + CStr(c) + ")" + "*" + CStr(d) + "=" + "24"
If a / (b - c) / d = 24 Then t = t + 1: s(t) = CStr(a) + "/" + "(" + CStr(b) + "-" + CStr(c) + ")" + "/" + CStr(d) + "=" + "24"
End If

'��������������ܴ�  so,����������������ܴ�
If (b - c + d) <> 0 Then
If a / (b - c + d) = 24 Then t = t + 1: s(t) = CStr(a) + "/" + "(" + CStr(b) + "-" + CStr(c) + "+" + CStr(d) + ")" + "=" + "24"
End If
If (b - c - d) <> 0 Then
If a / (b - c - d) = 24 Then t = t + 1: s(t) = CStr(a) + "/" + "(" + CStr(b) + "-" + CStr(c) + "-" + CStr(d) + ")" + "=" + "24"
End If
If (b + c + d) <> 0 Then
If a / (b + c + d) = 24 Then t = t + 1: s(t) = CStr(a) + "/" + "(" + CStr(b) + "+" + CStr(c) + "+" + CStr(d) + ")" + "=" + "24"
End If

If (a + b + c) * d = 24 Then t = t + 1: s(t) = "(" + CStr(a) + "+" + CStr(b) + "+" + CStr(c) + ")" + "*" + CStr(d) + "=" + "24"
If (a + b + c) / d = 24 Then t = t + 1: s(t) = "(" + CStr(a) + "+" + CStr(b) + "+" + CStr(c) + ")" + "/" + CStr(d) + "=" + "24"
If (a + b - c) * d = 24 Then t = t + 1: s(t) = "(" + CStr(a) + "+" + CStr(b) + "-" + CStr(c) + ")" + "*" + CStr(d) + "=" + "24"
If (a + b - c) / d = 24 Then t = t + 1: s(t) = "(" + CStr(a) + "+" + CStr(b) + "-" + CStr(c) + ")" + "/" + CStr(d) + "=" + "24"
If (a - b - c) * d = 24 Then t = t + 1: s(t) = "(" + CStr(a) + "-" + CStr(b) + "-" + CStr(c) + ")" + "*" + CStr(d) + "=" + "24"
If (a - b - c) / d = 24 Then t = t + 1: s(t) = "(" + CStr(a) + "-" + CStr(b) + "-" + CStr(c) + ")" + "/" + CStr(d) + "=" + "24"


If (a + b * c) * d = 24 Then t = t + 1: s(t) = "(" + CStr(a) + "+" + CStr(b) + "*" + CStr(c) + ")" + "*" + CStr(d) + "=" + "24"
If (a + b * c) / d = 24 Then t = t + 1: s(t) = "(" + CStr(a) + "+" + CStr(b) + "*" + CStr(c) + ")" + "/" + CStr(d) + "=" + "24"
If (a + b / c) * d = 24 Then t = t + 1: s(t) = "(" + CStr(a) + "+" + CStr(b) + "/" + CStr(c) + ")" + "*" + CStr(d) + "=" + "24"
If (a + b / c) / d = 24 Then t = t + 1: s(t) = "(" + CStr(a) + "+" + CStr(b) + "/" + CStr(c) + ")" + "/" + CStr(d) + "=" + "24"
If (a - b * c) * d = 24 Then t = t + 1: s(t) = "(" + CStr(a) + "-" + CStr(b) + "*" + CStr(c) + ")" + "*" + CStr(d) + "=" + "24"
If (a - b * c) / d = 24 Then t = t + 1: s(t) = "(" + CStr(a) + "-" + CStr(b) + "*" + CStr(c) + ")" + "/" + CStr(d) + "=" + "24"
If (a - b / c) * d = 24 Then t = t + 1: s(t) = "(" + CStr(a) + "-" + CStr(b) + "/" + CStr(c) + ")" + "*" + CStr(d) + "=" + "24"
If (a - b / c) / d = 24 Then t = t + 1: s(t) = "(" + CStr(a) + "-" + CStr(b) + "/" + CStr(c) + ")" + "/" + CStr(d) + "=" + "24"
If (a * b - c) * d = 24 Then t = t + 1: s(t) = "(" + CStr(a) + "*" + CStr(b) + "-" + CStr(c) + ")" + "*" + CStr(d) + "=" + "24"
If (a * b - c) / d = 24 Then t = t + 1: s(t) = "(" + CStr(a) + "*" + CStr(b) + "-" + CStr(c) + ")" + "/" + CStr(d) + "=" + "24"
If (a / b - c) * d = 24 Then t = t + 1: s(t) = "(" + CStr(a) + "/" + CStr(b) + "-" + CStr(c) + ")" + "*" + CStr(d) + "=" + "24"
If (a / b - c) / d = 24 Then t = t + 1: s(t) = "(" + CStr(a) + "/" + CStr(b) + "-" + CStr(c) + ")" + "/" + CStr(d) + "=" + "24"

If a / (b + c * d) = 24 Then t = t + 1: s(t) = CStr(a) + "/" + "(" + CStr(b) + "+" + CStr(c) + "*" + CStr(d) + ")" + "=" + "24"
If a / (b + c / d) = 24 Then t = t + 1: s(t) = CStr(a) + "/" + "(" + CStr(b) + "+" + CStr(c) + "/" + CStr(d) + ")" + "=" + "24"

If (b - c * d) <> 0 Then
If a / (b - c * d) = 24 Then t = t + 1: s(t) = CStr(a) + "/" + "(" + CStr(b) + "-" + CStr(c) + "*" + CStr(d) + ")" + "=" + "24"
End If

If (b - c / d) <> 0 Then
If a / (b - c / d) = 24 Then t = t + 1: s(t) = CStr(a) + "/" + "(" + CStr(b) + "-" + CStr(c) + "/" + CStr(d) + ")" + "=" + "24"
End If
If (b * c - d) <> 0 Then
If a / (b * c - d) = 24 Then t = t + 1: s(t) = CStr(a) + "/" + "(" + CStr(b) + "*" + CStr(c) + "-" + CStr(d) + ")" + "=" + "24"
End If
If (b / c - d) Then
If a / (b / c - d) = 24 Then t = t + 1: s(t) = CStr(a) + "/" + "(" + CStr(b) + "/" + CStr(c) + "-" + CStr(d) + ")" + "=" + "24"
End If

'˫����
If (c + d) <> 0 Then
If (a + b) * (c + d) = 24 Then t = t + 1: s(t) = "(" + CStr(a) + "+" + CStr(b) + ")" + "*" + "(" + CStr(c) + "+" + CStr(d) + ")" + "=" + "24"
If (a + b) / (c + d) = 24 Then t = t + 1: s(t) = "(" + CStr(a) + "+" + CStr(b) + ")" + "/" + "(" + CStr(c) + "+" + CStr(d) + ")" + "=" + "24"
If (a - b) / (c + d) = 24 Then t = t + 1: s(t) = "(" + CStr(a) + "-" + CStr(b) + ")" + "/" + "(" + CStr(c) + "+" + CStr(d) + ")" + "=" + "24"
End If
If (c - d) <> 0 Then
If (a + b) / (c - d) = 24 Then t = t + 1: s(t) = "(" + CStr(a) + "+" + CStr(b) + ")" + "/" + "(" + CStr(c) + "-" + CStr(d) + ")" + "=" + "24"
If (a + b) * (c - d) = 24 Then t = t + 1: s(t) = "(" + CStr(a) + "+" + CStr(b) + ")" + "*" + "(" + CStr(c) + "-" + CStr(d) + ")" + "=" + "24"
If (a - b) * (c - d) = 24 Then t = t + 1: s(t) = "(" + CStr(a) + "-" + CStr(b) + ")" + "*" + "(" + CStr(c) + "-" + CStr(d) + ")" + "=" + "24"
If (a - b) / (c - d) = 24 Then t = t + 1: s(t) = "(" + CStr(a) + "-" + CStr(b) + ")" + "/" + "(" + CStr(c) + "-" + CStr(d) + ")" + "=" + "24"
End If

End Sub




