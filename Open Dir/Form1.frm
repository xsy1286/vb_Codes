VERSION 5.00
Begin VB.Form Form1 
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   1320
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   1725
   Enabled         =   0   'False
   Icon            =   "Form1.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   1320
   ScaleWidth      =   1725
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  '����ȱʡ
   Visible         =   0   'False
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Load()
Shell "explorer D:\����\DSP", vbMaximizedFocus 'shell��������C��  Shell "explorer ·��"
Unload Me
End Sub
