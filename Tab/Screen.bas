Attribute VB_Name = "screen_attribute"
Public Declare Function GetSystemMetrics Lib "user32" (ByVal nIndex As Long) As Long
Public Const SM_CXSCREEN = 0
Public Const SM_CYSCREEN = 1

'Me.Width = GetSystemMetrics(SM_CXSCREEN) * Screen.TwipsPerPixelX 'screen.width�����ô����
'Me.Height = GetSystemMetrics(SM_CYSCREEN) * Screen.TwipsPerPixelY

'--------------------------------------------------------------------------------------------
Public Sub mid_Form(TheForm As Form)

TheForm.Left = (GetSystemMetrics(SM_CXSCREEN) * Screen.TwipsPerPixelX - TheForm.Width) / 2
TheForm.Top = (GetSystemMetrics(SM_CYSCREEN) * Screen.TwipsPerPixelY - TheForm.Height) / 2

End Sub
