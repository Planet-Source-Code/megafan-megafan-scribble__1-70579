VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ClsPile"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private Pile() As Long
Private PtrPile As Integer

Public Function InitPile(IntMax As Integer) As Boolean

    ReDim Pile(IntMax)
    PtrPile = 0

End Function

Public Sub Push(LngValue As Long)

    PtrPile = PtrPile + 1
    Pile(PtrPile) = LngValue
    
End Sub

Public Sub Pop(ByRef LngValue As Long)

    If PtrPile = 0 Then
        MsgBox "Pop impossible!"
    End If
    LngValue = Pile(PtrPile)
    PtrPile = PtrPile - 1

End Sub

