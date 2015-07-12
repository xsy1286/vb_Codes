VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "autoBuild"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
' The private collection object
Private m_ColvalueComs As Collection


Private Sub Class_Initialize()

    Set m_ColvalueComs = New Collection

End Sub


' This sub adds a new valueCom item to the collection.
Sub Add(ValItem As valueCom, Optional VarKey As Variant)

    'TODO: Initialize the new valueCom item's properties here

    m_ColvalueComs.Add ValItem, VarKey

End Sub


' This sub removes an item from the collection.
Sub Remove(VarIndex As Variant)

    m_ColvalueComs.Remove VarIndex

End Sub


' This function returns a valueCom item from the collection. It's the default method.
Function Item(VarIndex As Variant) As valueCom
Attribute Item.VB_UserMemId = 0

    Set Item = Nothing

    On Error Resume Next
    Set Item = m_ColvalueComs.Item(VarIndex)
    On Error GoTo 0

End Function


' This function returns a boolean value indicating whether the item exists in the collection or not.
Function Exists(VarIndex As Variant) As Boolean

    Dim objTest As valueCom
    Set objTest = Nothing

    If Count > 0 Then

        On Error Resume Next
        Set objTest = Item(VarIndex)
        On Error GoTo 0

    End If

    Exists = Not (objTest Is Nothing)

End Function


' This property returns the number of items in the collection.
Property Get Count() As Long

    Count = m_ColvalueComs.Count

End Property


' This sub removes all items from the collection.
Sub Clear()

    Set m_ColvalueComs = New Collection

End Sub


' This function adds "For Each" enumeration support. Must have a -4 DispID.
Function NewEnum() As IUnknown
Attribute NewEnum.VB_UserMemId = -4

    Set NewEnum = m_ColvalueComs.[_NewEnum]

End Function



