VERSION 5.00
Begin VB.Form Form3 
   Caption         =   "Tst"
   ClientHeight    =   3030
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   4560
   LinkTopic       =   "Form3"
   ScaleHeight     =   3030
   ScaleWidth      =   4560
   StartUpPosition =   3  '´°¿ÚÈ±Ê¡
   Begin VB.ListBox List1 
      Height          =   960
      Left            =   480
      TabIndex        =   1
      Top             =   1560
      Width           =   2415
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Command1"
      Height          =   855
      Left            =   720
      TabIndex        =   0
      Top             =   360
      Width           =   2415
   End
End
Attribute VB_Name = "Form3"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
List1.AddItem init_txt("TC", "1", "55")
List1.AddItem init_txt("TC", "2", "600")
List1.AddItem init_txt("TC", "mt1", "200")
List1.AddItem init_txt("TC", "mt2", "200")
List1.AddItem init_txt("TC", "mt3", "200")
End Sub

