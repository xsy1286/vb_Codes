VERSION 5.00
Begin VB.UserControl SingleValue 
   ClientHeight    =   1005
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   2730
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   ScaleHeight     =   1005
   ScaleWidth      =   2730
   Begin VB.VScrollBar scl2 
      Height          =   855
      Left            =   2400
      TabIndex        =   3
      Top             =   0
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.HScrollBar scl 
      Height          =   255
      Left            =   720
      TabIndex        =   2
      Top             =   600
      Width           =   1935
   End
   Begin VB.TextBox txt 
      Alignment       =   1  'Right Justify
      Height          =   285
      Left            =   720
      TabIndex        =   1
      Text            =   "Text1"
      Top             =   0
      Width           =   1935
   End
   Begin VB.Label lbl 
      BackStyle       =   0  'Transparent
      Caption         =   "Label1"
      Height          =   735
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   585
   End
End
Attribute VB_Name = "SingleValue"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Private Declare Function ShowWindow Lib "user32" (ByVal hwnd As Long, ByVal nCmdShow As Long) As Long
Private Declare Function GetParent Lib "user32" (ByVal hwnd As Long) As Long
Private Declare Function GetWindowRect Lib "user32" (ByVal hwnd As Long, lpRect As RECT) As Long
Private Type RECT
        Left As Long
         Top As Long
        Right As Long
        Bottom As Long
        fuck As Long
End Type
Private Declare Function MoveWindow Lib "user32" (ByVal hwnd As Long, ByVal X As Long, ByVal Y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal bRepaint As Long) As Long
Private Declare Function GetClientRect Lib "user32" (ByVal hwnd As Long, lpRect As RECT) As Long

Private retp As RECT
Private retm As RECT
'Property Set 可能有误，暂不使用
'Dim m_val As Integer
Private Const SW_HIDE = 0
Private Const SW_SHOW = 5

Private vis As Boolean

Public Enum HorV
 h = 0
 v = 1
End Enum
Private style As HorV
Public Event valChange(ByRef a As Integer)
Const str1 = "Out of Range"

'Public Enum backstyenum
' 透明 = 0
' 自设 = 1
'End Enum
Const banH = 10
Private Sub delayus(delaytime As Long)  '看电脑速度
Dim i As Long
For i = 1 To delaytime
DoEvents
Next i
End Sub

Public Property Get a_val() As Integer
 a_val = scl.Value
End Property
Public Property Let a_val(ByVal a As Integer)
    If a < scl.min Or a > scl.max Then MsgBox str1: Exit Property
    scl.Value = a
    scl2.Value = a
    PropertyChanged "val"
End Property
'Public Property Set a_val(ByVal a As Integer)      '传递引用
'    Set a = scl.Value
'End Property
Public Property Get min() As Integer
 min = scl.min
End Property
Public Property Let min(ByVal a As Integer)
    If a > scl.Value Then MsgBox str1:  Exit Property
    scl.min = a
    scl2.min = a
End Property
Public Property Get max() As Integer
 max = scl.max
End Property
Public Property Let max(ByVal a As Integer)
    If a < scl.Value Then MsgBox str1:  Exit Property
    scl.max = a
    scl2.max = a
End Property
Public Property Get stylef() As HorV
  stylef = style
End Property
Public Property Let stylef(ByVal a As HorV)
 style = a
 withStyle (style)
End Property



Private Sub scl_Change()
    txt.Text = CStr(scl.Value)
End Sub
Private Sub scl2_Change()
    txt.Text = CStr(scl2.Value)
End Sub

Private Sub txt_Change()
     Dim a As Integer
     a = Val(txt.Text)
     If a < scl.min Or a > scl.max Then
        MsgBox str1
       txt.Text = CStr(scl.Value)
        Exit Sub
      End If
    scl.Value = a
    scl2.Value = a
    RaiseEvent valChange(a)
End Sub



'***********************************UserControl Att.
Public Property Get m_color() As Long
  m_color = UserControl.BackColor
End Property

Public Property Let m_color(ByVal a As Long)
   UserControl.BackColor = a
End Property
Public Property Get BackStyle() As Integer
 BackStyle = UserControl.BackStyle
End Property
Public Property Let BackStyle(ByVal a As Integer)
  UserControl.BackStyle = a
End Property
Public Property Get mVisible() As Boolean
'UserControl.vi
  mVisible = vis
End Property
Public Property Let mVisible(ByVal a As Boolean)
  vis = a
If vis = True Then
 Call ShowWindow(UserControl.hwnd, SW_SHOW)
Else
 Call ShowWindow(UserControl.hwnd, SW_HIDE)
End If
 ' UserControl. = a
End Property
Public Property Get mLeft() As Integer

   Call GetClientRect(UserControl.hwnd, retm)
   mLeft = retm.Left * Screen.TwipsPerPixelX
End Property
Public Property Let mLeft(ByVal a As Integer)
'
'
'   Call MoveWindow(UserControl.hwnd, a / Screen.TwipsPerPixelX, retm.Top, retm.Right - retm.Left, retm.Bottom - retm.Top, 1)
'   Call GetClientRect(UserControl.hwnd, retm)


End Property

Public Property Get mTop() As Integer
   Call GetClientRect(UserControl.hwnd, retm)
   mTop = retm.Top * Screen.TwipsPerPixelY
End Property
Public Property Let mTop(ByVal a As Integer)
   Call GetClientRect(UserControl.hwnd, retm)
   Call MoveWindow(UserControl.hwnd, retm.Left, a / Screen.TwipsPerPixelY, retm.Right - retm.Left, retm.Bottom - retm.Top, 1)

   End Property
Public Property Get mWidth() As Integer
'      Call GetClientRect(UserControl.hwnd, retm)
'   mWidth = (retm.Right - retm.Left) * Screen.TwipsPerPixelX
End Property
Public Property Let mWidth(ByVal a As Integer)
   Call GetClientRect(UserControl.hwnd, retm)
   Call MoveWindow(UserControl.hwnd, retm.Left, retm.Top, a / Screen.TwipsPerPixelX, retm.Bottom - retm.Top, 1)

End Property
Public Property Get mHeight() As Integer
   Call GetClientRect(UserControl.hwnd, retm)
   mHeight = (retm.Bottom - retm.Top) * Screen.TwipsPerPixelY
End Property
Public Property Let mHeight(ByVal a As Integer)
'   Call GetClientRect(UserControl.hwnd, retm)
'   Call MoveWindow(UserControl.hwnd, retm.Left, retm.Top, retm.Right - retm.Left, a / Screen.TwipsPerPixelY, 1)
   
End Property

'***********************************UserControl Att. End


Public Property Get m_caption() As String
   m_caption = lbl.Caption
End Property
Public Property Let m_caption(ByVal a As String)
   lbl.Caption = a
End Property

'Public Property Set m_caption(ByVal a As String)      '传递引用
'    Set a = lbl.Caption
'End Property

Public Property Get lbl_BackStyle() As Integer
 lbl_BackStyle = lbl.BackStyle
End Property
Public Property Let lbl_BackStyle(ByVal a As Integer)
  lbl.BackStyle = a
End Property

Public Property Get lbl_color() As Long
    lbl_color = lbl.BackColor
End Property

Public Property Let lbl_color(ByVal a As Long)
    lbl.BackColor = a
End Property

Private Sub UserControl_Initialize()

    UserControl.BackStyle = 0
    lbl.BackStyle = 0
    With txt
        .Text = "Value"
        .Alignment = 1
    End With
     
End Sub

Private Sub UserControl_ReadProperties(PropBag As PropertyBag)

    With PropBag
        ' .ReadProperty
        UserControl.BackColor = .ReadProperty("color", &H0)
        scl.Value = .ReadProperty("val", 1)
        lbl.Caption = .ReadProperty("caption", "property")
        style = .ReadProperty("style", 0)
    End With

End Sub

Private Sub UserControl_WriteProperties(PropBag As PropertyBag)

   With PropBag
    .WriteProperty "val", scl.Value, 1
    .WriteProperty "caption", lbl.Caption, "property"
    .WriteProperty "color", UserControl.BackColor, &H0
    .WriteProperty "style", style, 0
   End With
   withStyle (style)
End Sub
Private Sub withStyle(a As HorV)
 If a = h Then
  toH
 Else
  toV
End If
End Sub

Private Sub toH()
  scl.Visible = True
  scl2.Visible = False
With UserControl
   .Height = 1005
   .Width = 2775
End With
 
  With scl
      .Height = 255
      .Left = 720
      .TabIndex = 2
      .Top = 400
      .Width = 1935
  End With
   With txt
     .Alignment = 1           'Right Justify
     .Height = 285
     .Left = 720
     .TabIndex = 1

      .Top = 0
      .Width = 1935
   End With
   With lbl

     .Height = 735
     .Left = 0
     .TabIndex = 0
      .Top = 0
      .Width = 585
   End With
   scl2.Visible = False
End Sub

Private Sub toV()
 
   scl2.Visible = True
   scl.Visible = False
  
    With UserControl
        .Height = 1320
        .Width = 2445
    End With

    With txt
        .Height = 375
        .Left = 240
        .TabIndex = 2
        .Top = 720
        .Width = 1455
    End With

    With scl2
        .Height = 975
        .Left = 1920
        .TabIndex = 0
        .Top = 120
        .Width = 230
    End With

    With lbl
        .Height = 195
        .Left = 240
        .TabIndex = 1
        .Top = 120
        .Width = 465
    End With

End Sub
