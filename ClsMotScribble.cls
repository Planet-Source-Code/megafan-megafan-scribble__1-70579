VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ClsMotScribble"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit


Private mMot As String
Private mValLettres As String
Private mPos As String
Private mPts As Integer

Property Let Mot(StrMot As String)

    mMot = StrMot

End Property

Property Get Mot() As String

    Mot = mMot

End Property

Property Let Pts(IntPts As Integer)

    mPts = IntPts

End Property

Property Get Pts() As Integer

    Pts = mPts

End Property

Property Let Pos(StrPos As String)

    mPos = StrPos

End Property

Property Get Pos() As String

    Pos = mPos

End Property

