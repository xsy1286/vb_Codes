VERSION 5.00
Begin VB.Form JustforMenu 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Form2"
   ClientHeight    =   3030
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   4560
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3030
   ScaleWidth      =   4560
   StartUpPosition =   3  '´°¿ÚÈ±Ê¡
   Begin VB.Menu n_pop 
      Caption         =   "0"
      Visible         =   0   'False
      Begin VB.Menu n_off 
         Caption         =   "off"
      End
   End
End
Attribute VB_Name = "JustforMenu"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Form_Load()

End Sub

Private Sub n_off_Click()
    Call UnHooK
    End
End Sub

Private Sub n_pop_Click()
  
End Sub
