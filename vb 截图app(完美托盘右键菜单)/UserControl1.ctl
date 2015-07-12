VERSION 5.00
Begin VB.UserControl UserControl1 
   ClientHeight    =   3600
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   4800
   ScaleHeight     =   3600
   ScaleWidth      =   4800
End
Attribute VB_Name = "UserControl1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Private Sub A()
Dir1.Left = Screen.Width / 2 - Dir1.Width / 2
Dir1.Top = Screen.Height / 2 - Dir1.Height / 2 - 600
Label1.Left = Dir1.Left - 450
Label1.Top = Dir1.Top - 1200
Text1.Left = Label1.Left + 1600
Text1.Top = Label1.Top
Command5.Left = Text1.Left + Text1.Width + 1100
Command5.Top = Text1.Top
Text2.Top = Dir1.Top - 600
Text2.Left = Dir1.Left - 1000
Combo1.Left = Text1.Left + Text1.Width + 100
Combo1.Top = Text1.Top
Check1.Left = Text2.Left + Text2.Width + 100
Check1.Top = Text2.Top
Text2.Text = Dir1.Path & "\" & Text1.Text & ".bmp"
Text1.Visible = True
Text2.Visible = True
Dir1.Visible = True
Label1.Visible = True
Combo1.Visible = True
Check1.Visible = True

End Sub

