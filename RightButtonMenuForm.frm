VERSION 5.00
Begin VB.Form rBtuM 
   Caption         =   "rightButtonMenu"
   ClientHeight    =   3030
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   4560
   LinkTopic       =   "Form2"
   ScaleHeight     =   3030
   ScaleWidth      =   4560
   StartUpPosition =   3  '´°¿ÚÈ±Ê¡
   Visible         =   0   'False
   Begin VB.Menu m1name 
      Caption         =   "m1"
      Visible         =   0   'False
      Begin VB.Menu offname 
         Caption         =   "off"
      End
      Begin VB.Menu editname 
         Caption         =   "eidt"
         Shortcut        =   ^Q
      End
   End
End
Attribute VB_Name = "rBtuM"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub editname_Click()
 frm1alltop = 0
EditForm.Show
End Sub

Private Sub m1name_Click()
'EditForm.show
End Sub

Private Sub offname_Click()
End
End Sub
