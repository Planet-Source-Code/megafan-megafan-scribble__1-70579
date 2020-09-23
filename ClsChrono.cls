VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ClsChrono"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Dim mDebut As Long
Dim mFin As Long
Dim mPas As Long
Dim mT0 As Long
Dim mValue As Long

Property Let Debut(LngDebut As Long)

   mDebut = LngDebut
   mValue = LngDebut
   
End Property

Property Let Fin(LngFin As Long)

   mFin = LngFin
   
End Property

Property Get Fin() As Long

   Fin = mFin
   
End Property

Property Let Pas(LngPas As Long)

   mPas = LngPas
   
End Property

Property Get Pas() As Long

    Pas = mPas
    
End Property

Property Let T0(LngT0 As Long)

   mT0 = LngT0
   
End Property

Property Get T0() As Long

   T0 = mT0

End Property

Property Let Value(LngValue As Long)
    
    mValue = LngValue

End Property
Property Get Value() As Long

   Value = mValue

End Property

