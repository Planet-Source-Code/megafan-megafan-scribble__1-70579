VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ClsDictionnaire"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private Dictionnaire() As String

Private Type TNoeud
    TabFils(26) As Long
    MotExite(26) As Boolean
End Type

Private Noeud() As TNoeud
Private PtrArbre As Long

Public Function InitDictionnaire(StrNom As String) As Boolean

    Dim TbTemp() As Byte
    Dim StrTemp As String
    Dim i As Long
    Dim Ptr As Long
    Dim Mot As String
    
    On Error GoTo GestErr

    FrmSplash.LblMessage.Caption = "Chargement du dictionnaire " & StrNom & " ..."
    DoEvents
    Debug.Print "Démarrage  : " & Time
       
    
    ' Initilisation des paramétres suivant le dictionnaire
    Select Case StrNom
        
            
        Case "ODS" ' ODS 5
            TbTemp = LoadResData(101, "DICO")
            MAXMOTS = 378989
            ReDim Noeud(722072)
            POINTSDELETTRE = Array(1, 3, 3, 2, 1, 4, 2, 4, 1, 8, 10, 1, 2, 1, 1, 3, 8, 1, 1, 1, 1, 4, 10, 10, 10, 10)
            STRPIOCHEDICO = "09A02B02C03D15E02F02G02H08I01J01K05L03M06N06O02P01Q06R06S06T06U02V01W01X01Y01Z02?"
            
        Case "TWL"
            TbTemp = LoadResData(102, "DICO")
            MAXMOTS = 178690
            ReDim Noeud(395182)
            POINTSDELETTRE = Array(1, 3, 3, 2, 1, 4, 2, 4, 1, 8, 5, 1, 3, 1, 1, 3, 10, 1, 1, 1, 1, 4, 4, 8, 4, 10)
            STRPIOCHEDICO = "09A02B02C04D12E02F03G02H09I01J01K04L02M06N08O02P01Q06R04S06T04U02V02W01X02Y01Z02?"
            
        Case "SOWPODS"
            TbTemp = LoadResData(103, "DICO")
            MAXMOTS = 216553
            ReDim Noeud(461495)
            POINTSDELETTRE = Array(1, 3, 3, 2, 1, 4, 2, 4, 1, 8, 5, 1, 3, 1, 1, 3, 10, 1, 1, 1, 1, 4, 4, 8, 4, 10)
            STRPIOCHEDICO = "09A02B02C04D12E02F03G02H09I01J01K04L02M06N08O02P01Q06R04S06T04U02V02W01X02Y01Z02?"
        
        Case "ZINGA"
            TbTemp = LoadResData(104, "DICO")
            MAXMOTS = 584982
            ReDim Noeud(1143408)
            POINTSDELETTRE = Array(1, 5, 2, 5, 1, 5, 8, 8, 1, 0, 0, 3, 3, 3, 1, 5, 10, 2, 2, 2, 3, 5, 0, 0, 0, 8)
            STRPIOCHEDICO = "14A03B06C03D11E03F02G02H12I00J00K05L05M05N15O03P01Q06R06S06T05U03V00W00X00Y02Z02?"
            ' Italian-language sets use these 120 tiles:
            ' 2 blank tiles (scoring 0 points)
            ' 1 point: O ×15, A ×14, I ×12, E ×11
            ' 2 points: C ×6, R ×6, S ×6, T ×6
            ' 3 points: L ×5, M ×5, N ×5, U ×5
            ' 5 points: B ×3, D ×3, F ×3, P ×3, V ×3
            ' 8 points: G ×2, H ×2, Z ×2
            ' 10 points: Q ×1
                
        Case Else
            MsgBox "Le dictionnaire '" & StrNom & "' n'est pas géré !" & vbCrLf & vbCrLf & "Le dictionnaire par défaut va être chargé.", vbCritical
            InitDictionnaire = False
            Exit Function
    End Select
    
    ' Chargement du dictionnaire depuis la ressource
    StrTemp = StrConv(TbTemp, vbUnicode)
    Dictionnaire = Split(StrTemp, vbCrLf)
    
    PtrArbre = 27
    FrmSplash.LblMessage.Caption = "Initialisation du dictionnaire " & StrNom & " ..."
    DoEvents
    Debug.Print "Fin chargement dictionnaire : " & Time
        
    For i = 1 To 26
        Noeud(0).TabFils(i) = i
    Next
    
    For i = 0 To UBound(Dictionnaire) - 1
        Mot = Dictionnaire(i)
        Ptr = Asc(Left(Mot, 1)) - 64
        AjouteArbre Ptr, Right(Mot, Len(Mot) - 1)
    Next
    
    Debug.Print "Fin création arbre: " & Time
    FrmSplash.LblMessage.Caption = "Dictionnaire " & StrNom & " initialisé ..."
    DoEvents
    InitDictionnaire = True
    
    Exit Function

GestErr:
    Resume Fin
Fin:
    InitDictionnaire = False
    
End Function

Private Sub AjouteArbre(LngPtr As Long, LettresRestantes As String)

    Dim Ptr As Long
    Dim Fils As Integer
    
    Fils = Asc(Left(LettresRestantes, 1)) - 64
    Ptr = Noeud(LngPtr).TabFils(Fils)
    
    If Ptr = 0 Then
        PtrArbre = PtrArbre + 1
        Ptr = PtrArbre
    End If
   
    Noeud(LngPtr).TabFils(Fils) = Ptr
        
    LettresRestantes = Right(LettresRestantes, Len(LettresRestantes) - 1)
    If LettresRestantes <> "" Then
        AjouteArbre Ptr, LettresRestantes
    Else
        Noeud(LngPtr).MotExite(Fils) = True
    End If
    
End Sub

Public Function MotExisteByNoeud(LngPere, IntIdx As Integer) As Boolean
    
    If Noeud(LngPere).MotExite(IntIdx) Then MotExisteByNoeud = True
    
End Function

Public Function GetNoeud(LngPere, IntIdx As Integer) As Long
        
    GetNoeud = Noeud(LngPere).TabFils(IntIdx)

End Function
