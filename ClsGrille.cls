VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ClsGrille"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private Grille(14, 14) As String
Private GrillePoints(14, 14) As String
Private LastCol As Integer
Private LastLig As Integer

Private Ana As ClsMots
Private mPremierMot As Boolean
Private mOrientation As String
Private mSolutionTrouvee As Boolean
Private mAbort As Boolean
Private mIsSearching As Boolean

Private PtrSolution As Long
Private TabSolution() As ClsMotScribble


Public Function ChercheLettre(Optional BlnDebut As Boolean) As String
    
    Dim i As Integer
    Dim j As Integer
       
    If BlnDebut = True Then
        LastCol = 0
        LastLig = 0
        Exit Function
    End If
    
    Do
        If LastCol < 15 Then
            LastCol = LastCol + 1
            If Grille(LastCol - 1, LastLig) <> "" Then
                ChercheLettre = Chr(65 + LastLig) & CStr(LastCol)
                Exit Function
            End If
        Else
            LastCol = 0
            LastLig = LastLig + 1
            If LastLig > 14 Then
                ChercheLettre = ""
                Exit Function
            End If
        End If
    Loop While 1

End Function

Public Sub SwapGrille()

    Dim i As Integer
    Dim j As Integer
    Dim Pos As String
    Dim GrilleTemp(15, 15) As String
            
    DoEvents
    For j = 0 To 14
        For i = 0 To 14
            GrilleTemp(i, j) = Grille(j, i)
        Next
    Next
    
    DoEvents
    For j = 0 To 14
        For i = 0 To 14
            Grille(i, j) = GrilleTemp(i, j)
        Next
    Next
    
    DoEvents
    If mOrientation = "H" Then
        mOrientation = "V"
    Else
        mOrientation = "H"
    End If
End Sub

Private Sub Class_Initialize()

    Dim StrTemp(15) As String
    Dim i As Integer
    Dim j As Integer

    StrTemp(1) = "400100040001004"
    StrTemp(2) = "030002000200030"
    StrTemp(3) = "003000101000300"
    StrTemp(4) = "100300010003001"
    StrTemp(5) = "000030000030000"
    StrTemp(6) = "020002000200020"
    StrTemp(7) = "001000101000100"
    StrTemp(8) = "400100030001004"
    StrTemp(9) = "001000101000100"
    StrTemp(10) = "020002000200020"
    StrTemp(11) = "000030000030000"
    StrTemp(12) = "100300010003001"
    StrTemp(13) = "003000101000300"
    StrTemp(14) = "030002000200030"
    StrTemp(15) = "400100040001004"
    
    For j = 0 To 14
        For i = 0 To 14
            GrillePoints(i, j) = Mid(StrTemp(j + 1), i + 1, 1)
        Next
    Next
    
    mOrientation = "H"
    
    Set Ana = New ClsMots
    Ana.Dictionnaire = Dico
    
    mPremierMot = True

End Sub

Property Get IsPremierMot() As Boolean

    IsPremierMot = mPremierMot
    
End Property

Property Let IsPremierMot(BlnPM As Boolean)

    mPremierMot = BlnPM
    
End Property

Property Get Cell(x As Integer, y As Integer) As String

    Cell = Grille(x, y)
    

End Property

Property Let Cell(x As Integer, y As Integer, StrValeur As String)

    Grille(x, y) = StrValeur

End Property

Public Function PlacerMot(ms As ClsMotScribble) As Boolean

    Dim col As Integer
    Dim lig As Integer
    Dim i As Integer
    Dim Orientation As Integer '1 =H 2=V
    Dim StrPosition As String
    Dim StrMot As String
    
    StrMot = ms.Mot
    StrPosition = ms.Pos

    PlacerMot = False
    StrPosition = Trim(UCase(StrPosition))
    If Len(StrPosition) > 3 Then Exit Function
    
    If Asc(Left(StrPosition, 1)) >= Asc("A") And Asc(Left(StrPosition, 1)) <= Asc("O") Then
        Orientation = 1
        lig = Asc(Left(StrPosition, 1)) - 65
        StrPosition = Right(StrPosition, Len(StrPosition) - 1)
        If Val(StrPosition) < 1 Or Val(StrPosition) > 15 Then Exit Function
        col = Val(StrPosition) - 1
        
    Else
        Orientation = 2
        If Asc(Right(StrPosition, 1)) < Asc("A") And Asc(Right(StrPosition, 1)) > Asc("O") Then Exit Function
        lig = Asc(Right(StrPosition, 1)) - 65
        StrPosition = Left(StrPosition, Len(StrPosition) - 1)
        If Val(StrPosition) < 1 Or Val(StrPosition) > 15 Then Exit Function
        col = Val(StrPosition) - 1
    End If
    
    
    If Orientation = 1 Then
        If col + Len(StrMot) > 15 Then Exit Function
        For i = 1 To Len(StrMot)
            If Mid(StrMot, i, 1) = "-" Then
                Grille(col + i - 1, lig) = ""
            Else
                Grille(col + i - 1, lig) = Mid(StrMot, i, 1)
            End If
        Next
    Else
        If lig + Len(StrMot) > 15 Then Exit Function
        For i = 1 To Len(StrMot)
            If Mid(StrMot, i, 1) = "-" Then
                Grille(col, lig + i - 1) = ""
            Else
                Grille(col, lig + i - 1) = Mid(StrMot, i, 1)
            End If
        Next
    End If

End Function

Public Function GetLigne(StrPosition As String) As Integer
    
    GetLigne = -1

    If Asc(Left(StrPosition, 1)) >= Asc("A") And Asc(Left(StrPosition, 1)) <= Asc("O") Then
        GetLigne = Asc(Left(StrPosition, 1)) - 65
    Else
        If Asc(Right(StrPosition, 1)) < Asc("A") And Asc(Right(StrPosition, 1)) > Asc("O") Then Exit Function
        GetLigne = Asc(Right(StrPosition, 1)) - 65
    End If
        
End Function

Public Function GetColonne(StrPosition As String) As Integer

    GetColonne = -1

    If Asc(Left(StrPosition, 1)) >= Asc("A") And Asc(Left(StrPosition, 1)) <= Asc("O") Then
        If Val(Right(StrPosition, Len(StrPosition) - 1)) < 1 Or Val(Right(StrPosition, Len(StrPosition) - 1)) > 15 Then Exit Function
        GetColonne = Val(Right(StrPosition, Len(StrPosition) - 1)) - 1
    Else
        If Asc(Right(StrPosition, 1)) < Asc("A") And Asc(Right(StrPosition, 1)) > Asc("O") Then Exit Function
        If Val(Left(StrPosition, Len(StrPosition) - 1)) < 1 Or Val(Left(StrPosition, Len(StrPosition) - 1)) > 15 Then Exit Function
        GetColonne = Val(Left(StrPosition, Len(StrPosition) - 1)) - 1
    End If

End Function

Public Function GetMot(StrPos As String) As String

    Dim col As Integer
    Dim lig As Integer
    Dim i As Integer
    Dim Orientation As Integer '1 =H 2=V
    Dim StrPosition As String
    
    StrPosition = Trim(UCase(StrPos))
    If Len(StrPosition) > 3 Then Exit Function
    
    If Asc(Left(StrPosition, 1)) >= Asc("A") And Asc(Left(StrPosition, 1)) <= Asc("O") Then
        Orientation = 1
        lig = Asc(Left(StrPosition, 1)) - 65
        StrPosition = Right(StrPosition, Len(StrPosition) - 1)
        'If Val(StrPosition) < 1 Or Val(StrPosition) > 15 Then Exit Function
        col = Val(StrPosition) - 1
        
    Else
        Orientation = 2
        If Asc(Right(StrPosition, 1)) < Asc("A") And Asc(Right(StrPosition, 1)) > Asc("O") Then Exit Function
        lig = Asc(Right(StrPosition, 1)) - 65
        StrPosition = Left(StrPosition, Len(StrPosition) - 1)
        'If Val(StrPosition) < 1 Or Val(StrPosition) > 15 Then Exit Function
        col = Val(StrPosition) - 1
    End If
    
    
    If Orientation = 1 Then
        i = 0
        Do
            i = i + 1
            If Grille(col + i - 1, lig) <> "" Then
                GetMot = GetMot + Grille(col + i - 1, lig)
            Else
                Exit Do
            End If
        Loop While (col + i) < 15
    Else
        i = 0
        Do
            i = i + 1
            If Grille(col, lig + i - 1) <> "" Then
                GetMot = GetMot + Grille(col, lig + i - 1)
            Else
                Exit Do
            End If
        Loop While (lig + i) < 15
    End If

    GetMot = UCase(GetMot)

End Function

Property Get GrPoints(x As Integer, y As Integer) As Integer

    GrPoints = GrillePoints(x, y)
    
End Property

Public Function ComptePointsMotHorizontal(StrMot As String, StrPosition As String, IntOrientation As Integer) As ClsMotScribble

    Dim Car As String
    Dim i As Integer
    Dim j As Integer
    Dim Pos As Integer
    Dim Pts As Integer
    Dim StrLocalMot As String
    Dim StrTempMot As String
    Dim StrTempMot2 As String
    Dim StrJocker As String
        
    Set ComptePointsMotHorizontal = New ClsMotScribble
    
    i = InStr(StrMot, " ")
    
    If i = 0 Then
        '  0 jockers
        Pts = ComptePointsMotHorizontalP(StrMot, StrPosition)
        ComptePointsMotHorizontal.Mot = StrMot
        ComptePointsMotHorizontal.Pts = Pts
        If IntOrientation = 0 Then
            ComptePointsMotHorizontal.Pos = StrPosition
        Else
            ComptePointsMotHorizontal.Pos = InversePos(StrPosition)
        End If
        Set ComptePointsMotHorizontal = ComptePointsMotHorizontal
        Exit Function
    End If
    
    ' Enlever '(' et ')' entourants le(s) jockers
    StrLocalMot = Left(StrMot, i - 1)
    StrJocker = Right(StrMot, Len(StrMot) - i - 1)
    StrJocker = Left(StrJocker, Len(StrJocker) - 1)
    
    If Len(StrJocker) = 1 Then
        ' 1 Jocker
        For i = 1 To Len(StrLocalMot)
            StrTempMot = StrLocalMot
            Car = Left(StrJocker, 1)
            If Mid(StrLocalMot, i, 1) = Car Then
                ' Sauf si case déja utilisée....
                If Cell(GetColonne(StrPosition) + i - 1, GetLigne(StrPosition)) = "" Then
                    Mid(StrTempMot, i, 1) = UCase(Car)
                    Pts = ComptePointsMotHorizontalP(StrTempMot, StrPosition)
                    If Pts >= ComptePointsMotHorizontal.Pts Then
                        ComptePointsMotHorizontal.Pts = Pts
                        ComptePointsMotHorizontal.Mot = StrTempMot
                        If IntOrientation = 0 Then
                            ComptePointsMotHorizontal.Pos = StrPosition
                        Else
                            ComptePointsMotHorizontal.Pos = InversePos(StrPosition)
                        End If
                        Set ComptePointsMotHorizontal = ComptePointsMotHorizontal
                    End If
                End If
            End If
        Next
        Exit Function
    End If
    
    ' 2 Jockers
    Dim Jock1 As String
    Dim Jock2 As String
    
    Jock1 = Left(StrJocker, 1)
    Jock2 = Right(StrJocker, 1)
    
    ' boucle pour le premier jocker
    For i = 1 To Len(StrLocalMot)
        StrTempMot = StrLocalMot
        If Mid(StrLocalMot, i, 1) = Jock1 Then
            If Cell(GetColonne(StrPosition) + i - 1, GetLigne(StrPosition)) = "" Then
                ' on peut mettre le jock 1
                Mid(StrTempMot, i) = UCase(Jock1)
                ' boucle pour le deuxieme jocker
                For j = 1 To Len(StrLocalMot)
                    StrTempMot2 = StrTempMot
                    If Mid(StrTempMot, j, 1) = Jock2 Then
                        If Cell(GetColonne(StrPosition) + j - 1, GetLigne(StrPosition)) = "" Then
                            ' on peut mettre le jock 2
                            Mid(StrTempMot2, j) = UCase(Jock2)
                            Pts = ComptePointsMotHorizontalP(StrTempMot, StrPosition)
                            If Pts >= ComptePointsMotHorizontal.Pts Then
                                ComptePointsMotHorizontal.Pts = Pts
                                ComptePointsMotHorizontal.Mot = StrTempMot2
                                If IntOrientation = 0 Then
                                    ComptePointsMotHorizontal.Pos = StrPosition
                                Else
                                    ComptePointsMotHorizontal.Pos = InversePos(StrPosition)
                                End If
                                Set ComptePointsMotHorizontal = ComptePointsMotHorizontal
                            End If
                        End If
                    End If
                Next
            End If
        End If
    Next
    
End Function

Private Function StrPoints(StrTemp As String)

    Dim i As Integer
    
    For i = 1 To Len(StrTemp)
        StrPoints = StrPoints & Hex(PointLettres(Mid(StrTemp, i, 1)))
    Next

End Function

Private Function ComptePointsMotHorizontalP(ByRef StrMot As String, StrPosition As String) As Integer

    Dim lig As Integer
    Dim col As Integer
    Dim Car As String
    Dim i As Integer
    Dim j As Integer
    Dim l As Integer
    Dim Bonus As Integer
    Dim PointLettre As Integer
    Dim Multiplicateur As Integer
    Dim DebutMot As Integer
    Dim FinMot As Integer
    Dim TempGrille(14, 14) As String
    Dim PointsMotVertical As Integer
       
    For j = 0 To 14
        For i = 0 To 14
            TempGrille(i, j) = Grille(i, j)
        Next
    Next
    DoEvents

    lig = GetLigne(StrPosition)
    col = GetColonne(StrPosition)
        
    l = 0
    For i = 1 To Len(StrMot)
        If TempGrille(i + col - 1, lig) = "" Then l = l + 1
        If Grille(i + col - 1, lig) = "" Then
            Car = Mid(StrMot, i, 1)
        Else
            Car = Grille(i + col - 1, lig)
            ' Forcer l'ecriture d'un jocker dans Strmot.
            Mid(StrMot, i, 1) = Car
        End If
        TempGrille(i + col - 1, lig) = Car
    Next
    
    
    ComptePointsMotHorizontalP = 0
    If l = 7 Then Bonus = 50
       
    Multiplicateur = 1
       
    For i = col To 14
        Car = TempGrille(i, lig)
        If Car = "" Then Exit For
        PointLettre = PointLettres(Car)
        If TempGrille(i, lig) <> Grille(i, lig) Then
            Select Case GrillePoints(i, lig)
                Case 1
                    PointLettre = 2 * PointLettre
                Case 2
                    PointLettre = 3 * PointLettre
                Case 3
                    Multiplicateur = Multiplicateur * 2
                Case 4
                    Multiplicateur = Multiplicateur * 3
            End Select
        End If
        ComptePointsMotHorizontalP = ComptePointsMotHorizontalP + PointLettre
    Next
    
    ComptePointsMotHorizontalP = ComptePointsMotHorizontalP * Multiplicateur
    
    
    ' Les mots verticaux crées sont comptés
    For i = col To 14
        If (Grille(i, lig) <> TempGrille(i, lig)) And TempGrille(i, lig) <> "?" Then
            Multiplicateur = 1
            PointsMotVertical = 0
            DebutMot = lig
            If lig > 0 Then
                Do
                    Car = TempGrille(i, DebutMot - 1)
                    If Car = "" Then Exit Do
                    DebutMot = DebutMot - 1
                Loop While DebutMot > 0
            End If
            FinMot = lig
            If lig < 14 Then
                Do
                    Car = TempGrille(i, FinMot + 1)
                    If Car = "" Then Exit Do
                    FinMot = FinMot + 1
                Loop While FinMot < 14
            End If
                
            If FinMot - DebutMot >= 1 Then
                For j = DebutMot To FinMot
                    Car = TempGrille(i, j)
                                       
                    ' Compter les bonus sur les nouvelles lettre
                    PointLettre = PointLettres(Car)
                    If Grille(i, j) = "" Then
                        Select Case GrillePoints(i, lig)
                            Case 1
                                PointLettre = 2 * PointLettre
                            Case 2
                                PointLettre = 3 * PointLettre
                            Case 3
                                Multiplicateur = Multiplicateur * 2
                            Case 4
                                Multiplicateur = Multiplicateur * 3
                        End Select
                    End If
                    PointsMotVertical = PointsMotVertical + PointLettre
                Next
                ComptePointsMotHorizontalP = ComptePointsMotHorizontalP + PointsMotVertical * Multiplicateur
            End If
        End If
     Next
    
    ComptePointsMotHorizontalP = ComptePointsMotHorizontalP + Bonus

End Function

Public Function InversePos(StrPosition As String) As String

    Dim col As Integer
    Dim lig As Integer
    
    lig = GetLigne(StrPosition)
    col = GetColonne(StrPosition)
    
    If Asc(Left(StrPosition, 1)) >= Asc("A") And Asc(Left(StrPosition, 1)) <= Asc("O") Then
        ' Horizontal
        InversePos = Format(lig + 1) & Chr(65 + col)
    Else
        ' Verticale
        InversePos = Chr(65 + col) & Format(lig + 1)
    End If


End Function

Public Function PointLettres(StrCar As String) As Integer

    If Asc(StrCar) < 97 Then
        PointLettres = 0
    Else
        PointLettres = POINTSDELETTRE(Asc(StrCar) - 97)
    End If
    
End Function

Private Sub ComputeZoneHaut(StrPosition As String, StrMaPioche As String, IntOrientation As Integer)

    Dim lig As Integer
    Dim col As Integer
    Dim i As Integer
    Dim IntStop As Integer
    Dim StrTypeCell As String
    Dim TabPossibilites(15) As String
    Dim Debut As Integer
    Dim Fin As Integer
    Dim ParcoursTab As Integer
    Dim StrTemp As String
    Dim NbMots As Long
    Dim LgMini As Integer
    Dim LgMaxi As Integer
    Dim k As Long
    Dim StrMot As String
    Dim l As Integer
    Dim BlnOk As Boolean
    Dim StrEntete As String
    Dim StrSuffixe As String
    Dim MsTemp As ClsMotScribble
    Dim StrVraiMot As String
        
    lig = GetLigne(StrPosition)
    col = GetColonne(StrPosition)
            
    ' Déterminer s'il y a une zone Haute (possibilité d'ajouter une lettre au début d'un mot vertical )
    If lig <> 0 Then
        If Cell(col, lig - 1) = "" Then
            
            '////////////////////////////////////////////////////////////////////////////////
            
            Call GetTabPossibiliteFrom(col, lig - 1, Len(StrMaPioche), 0, TabPossibilites())
                        
            For i = 0 To 14
                If TabPossibilites(i) <> "" Then
                    Debut = i
                    Exit For
                End If
            Next
            
            For i = Debut To 14
                If TabPossibilites(i) = "" Or i = 14 Then
                    Fin = i
                    Exit For
                End If
            Next
            
            For ParcoursTab = Debut To col
                
                StrTemp = ""
                
                IntStop = Len(StrMaPioche)
                For i = ParcoursTab To Fin
                    DoEvents
                    If TabPossibilites(i) = "F" Then
                        StrTemp = StrTemp + Cell(i, lig - 1)
                    Else
                        StrTemp = StrTemp + "*"
                        IntStop = IntStop - 1
                        If IntStop = 0 Then Exit For
                    End If
                Next
                LgMini = col - ParcoursTab + 1
                LgMaxi = Fin - ParcoursTab + 1
                
                NbMots = Ana.AnaGrammSelonMasque(StrTemp, StrMaPioche, LgMini)
                                
                For k = 1 To NbMots
                    StrVraiMot = Ana.GetMot(k)
                    DoEvents
                    '************************************
                    ' TODO
                    ' TODO
                    ' TODO
                    'VerifierNecessiteTousJocker StrVraiMot, TxtLettresUC
                    '*************************************
                    StrMot = SansJock(StrVraiMot)
                    
                    If Len(StrMot) <= LgMaxi Then
                        
                        BlnOk = True
                        '/////////////////////////////////////////////////////////////////////
                        'Optimisation
                        l = 1
                        Do
                            If TabPossibilites(ParcoursTab + l - 1) <> "L" Then Exit Do
                            l = l + 1
                            'if l>len(strmot) then msgbox "Impossible ???"
                        Loop While 1
                        '/////////////////////////////////////////////////////////////////////
                        
                        For l = l To Len(StrMot)
                            ' B
                            If InStr(TabPossibilites(ParcoursTab + l - 1), "B") Then
                                Cell(ParcoursTab + l - 1, lig - 1) = Mid(StrMot, l, 1)
                                ' le mot commence peut etre en lig-n
                                If lig - 2 >= 0 Then
                                    For i = lig - 1 To 0 Step -1
                                        DoEvents
                                        If Cell(ParcoursTab + l - 1, i) = "" Then Exit For
                                    Next
                                Else
                                    i = -1
                                End If
                                
                                If Ana.IsMotexiste(GetMot(Format(ParcoursTab + l) & Chr(i + 65 + 1))) = False Then
                                    Cell(ParcoursTab + l - 1, lig - 1) = ""
                                    BlnOk = False
                                    Exit For
                                End If
                                Cell(ParcoursTab + l - 1, lig - 1) = ""
                            End If
                            
                            ' H
                            If InStr(TabPossibilites(ParcoursTab + l - 1), "H") Then
                                Cell(ParcoursTab + l - 1, lig - 1) = Mid(StrMot, l, 1)
                                ' le mot commence en lig-n
                                If lig - 2 >= 0 Then
                                    For i = lig - 1 To 0 Step -1
                                        If Cell(ParcoursTab + l - 1, i) = "" Then Exit For
                                    Next
                                Else
                                    i = -1
                                End If
                                If Ana.IsMotexiste(GetMot(Format(ParcoursTab + l) & Chr(i + 65 + 1))) = False Then
                                    Cell(ParcoursTab + l - 1, lig - 1) = ""
                                    BlnOk = False
                                    Exit For
                                End If
                                Cell(ParcoursTab + l - 1, lig - 1) = ""
                            
                            End If
                        Next
                        
                        ' Un mot est correct, mais verifier s'il y a des lettres fixes avant le mot ou après
                        If BlnOk Then
                            ' En-tete
                            StrEntete = ""
                            If ParcoursTab <> 0 Then
                                For i = ParcoursTab - 1 To 0 Step -1
                                    If TabPossibilites(i) <> "F" Then Exit For
                                    StrEntete = Cell(i, lig - 1) & StrEntete
                                Next
                            End If
                        
                            ' Suffixe
                            StrSuffixe = ""
                            For i = ParcoursTab + Len(StrMot) To 14
                                If TabPossibilites(i) <> "F" Then Exit For
                                StrSuffixe = StrSuffixe + Cell(i, lig - 1)
                            Next
                        
                            If Ana.IsMotexiste(StrEntete & StrMot & StrSuffixe) = False Then BlnOk = False
                        End If
                        
                        If BlnOk Then
                            Dim StrToAff As String
                            Dim StrP As String
                            Dim points As Integer
                            StrToAff = ""
                            If StrEntete <> "" Then StrToAff = StrEntete
                            StrToAff = StrToAff & StrMot
                            If StrSuffixe <> "" Then StrToAff = StrToAff & StrSuffixe
                            StrToAff = LCase(StrToAff & AjouteJocker(StrVraiMot))
                            
                            StrP = Chr(65 + lig - 1) & Format(ParcoursTab + 1)
                            Set TabSolution(PtrSolution) = ComptePointsMotHorizontal(StrToAff, StrP, IntOrientation)
                            PtrSolution = PtrSolution + 1
                        End If
                    End If
                Next
                
                ' Optimisation
                Do
                    If TabPossibilites(ParcoursTab) = "F" Then
                        ParcoursTab = ParcoursTab + 1
                        If ParcoursTab > col Then Exit For
                    Else
                        Exit Do
                    End If
                Loop While 1
            Next
            
        End If
    End If

End Sub

Private Sub ComputeZoneCentrale(StrPosition As String, StrMaPioche As String, IntOrientation As Integer)
    
    Dim lig As Integer
    Dim col As Integer
    Dim i As Integer
    Dim IntStop As Integer
    Dim StrTypeCell As String
    Dim TabPossibilites(15) As String
    Dim Debut As Integer
    Dim Fin As Integer
    Dim ParcoursTab As Integer
    Dim NbMots As Long
    Dim LgMini As Integer
    Dim LgMaxi As Integer
    Dim StrTemp As String
    Dim k As Long
    Dim StrMot As String
    Dim l As Integer
    Dim BlnOk As Boolean
    Dim StrEntete As String
    Dim StrSuffixe As String
    Dim MsTemp As ClsMotScribble
    Dim StrVraiMot As String
    
    Static OptimisationCol As Integer
    Static OptimisationLig  As Integer
    Static OptimisationOK As Boolean
        
    lig = GetLigne(StrPosition)
    col = GetColonne(StrPosition)
    
    
    ' Déterminer la zone Centrale (possibilité d'ajouter une lettre à droite ou à gauche)
    
    '////////////////////////////////////////////////////////////////////////////
    
    Call GetTabPossibiliteFrom(col - 1, lig, Len(StrMaPioche) + 1, 1, TabPossibilites())
    TabPossibilites(col) = "F"
    
    If OptimisationOK = True Then
        If col = OptimisationCol + 1 And OptimisationLig = lig Then
            OptimisationCol = col
            Exit Sub
        Else
            OptimisationOK = False
        End If
    End If
        
    For i = 0 To 14
        If TabPossibilites(i) <> "" Then
            Debut = i
            Exit For
        End If
    Next
        
    For i = Debut To 14
        If TabPossibilites(i) = "" Or i = 14 Then
            Fin = i
            Exit For
        End If
    Next
    
    For ParcoursTab = Debut To col
    
        StrTemp = ""
        
        BlnOk = False
                    
        IntStop = Len(StrMaPioche)
        For i = ParcoursTab To Fin
            DoEvents
            If TabPossibilites(i) = "F" Then
                StrTemp = StrTemp & Cell(i, lig)
            Else
                StrTemp = StrTemp + "*"
                BlnOk = True
                IntStop = IntStop - 1
                If IntStop = 0 Then Exit For
            End If
        Next
        LgMini = col - ParcoursTab
        
        
        If LgMini = 0 Then
            LgMini = 1
            For i = 0 To 14
                If TabPossibilites(ParcoursTab + i) = "F" Then
                    LgMini = LgMini + 1
                Else
                    Exit For
                End If
            Next
        End If
        
        If Len(StrTemp) = 1 Then StrTemp = StrTemp + Cell(col, lig)
        
        LgMaxi = Fin - ParcoursTab + 2
        
        If BlnOk Then
            
            NbMots = Ana.AnaGrammSelonMasque(StrTemp, StrMaPioche, LgMini)
            
            For k = 1 To NbMots
                StrVraiMot = Ana.GetMot(k)
                StrMot = SansJock(StrVraiMot)
                DoEvents
                                
                If Len(StrMot) <= LgMaxi Then
                
                    BlnOk = True
                    '/////////////////////////////////////////////////////////////////////
                    'Optimisation
                    l = 1
                    Do
                        If TabPossibilites(ParcoursTab + l - 1) <> "L" Then Exit Do
                        l = l + 1
                        'if l>len(strmot) then msgbox "Impossible ???"
                    Loop While 1
                    '/////////////////////////////////////////////////////////////////////
                    
                    For l = l To Len(StrMot)
                        ' B
                        If InStr(TabPossibilites(ParcoursTab + l - 1), "B") Then
                            Cell(ParcoursTab + l - 1, lig) = Mid(StrMot, l, 1)
                            ' le mot commence peut etre en lig-n
                            If lig > 0 Then
                                For i = lig - 1 To 0 Step -1
                                    If Cell(ParcoursTab + l - 1, i) = "" Then Exit For
                                Next
                            Else
                                i = -1
                            End If
                            
                            If Ana.IsMotexiste(GetMot(Format(ParcoursTab + l) & Chr(i + 65 + 1))) = False Then
                                Cell(ParcoursTab + l - 1, lig) = ""
                                BlnOk = False
                                Exit For
                            End If
                            Cell(ParcoursTab + l - 1, lig) = ""
                        End If
                        
                        ' H
                        If InStr(TabPossibilites(ParcoursTab + l - 1), "H") Then
                            Cell(ParcoursTab + l - 1, lig) = Mid(StrMot, l, 1)
                            ' le mot commence en lig-n
                            If lig > 1 Then
                                For i = lig - 1 To 0 Step -1
                                    If Cell(ParcoursTab + l - 1, i) = "" Then Exit For
                                Next
                            Else
                                i = -1
                            End If
                            If Ana.IsMotexiste(GetMot(Format(ParcoursTab + l) & Chr(i + 65 + 1))) = False Then
                                Cell(ParcoursTab + l - 1, lig) = ""
                                BlnOk = False
                                Exit For
                            End If
                            Cell(ParcoursTab + l - 1, lig) = ""
                        
                        End If
                        
                    Next
                        
                    ' Un mot est correct, mais verifier s'il y a des lettres fixes avant le mot ou après
                    If BlnOk Then
                        ' En-tete
                        StrEntete = ""
                        If ParcoursTab <> 0 Then
                            For i = ParcoursTab - 1 To 0 Step -1
                                If TabPossibilites(i) <> "F" Then Exit For
                                StrEntete = Cell(i, lig) & StrEntete
                            Next
                        End If
                        
                        ' Suffixe
                        StrSuffixe = ""
                        For i = ParcoursTab + Len(StrMot) To 14
                            If TabPossibilites(i) <> "F" Then Exit For
                            StrSuffixe = StrSuffixe + Cell(i, lig)
                        Next
                        
                        If Ana.IsMotexiste(StrEntete & StrMot & StrSuffixe) = False Then BlnOk = False
                    End If
                    
                    If BlnOk Then
                        Dim StrToAff As String
                        Dim StrP As String
                        Dim points As Integer
                        StrToAff = ""
                        If StrEntete <> "" Then StrToAff = StrEntete
                        StrToAff = StrToAff & StrMot
                        If StrSuffixe <> "" Then StrToAff = StrToAff & StrSuffixe
                        StrToAff = LCase(StrToAff & AjouteJocker(StrVraiMot))
                        
                        StrP = Chr(65 + lig) & Format(ParcoursTab + 1)
                        Set TabSolution(PtrSolution) = ComptePointsMotHorizontal(StrToAff, StrP, IntOrientation)
                        PtrSolution = PtrSolution + 1
                    End If
    
                End If
            Next k
            
            ' Optimisation
            Do
                If TabPossibilites(ParcoursTab) = "F" Then
                    ParcoursTab = ParcoursTab + 1
                    If ParcoursTab > col Then Exit For
                Else
                    Exit Do
                End If
            Loop While 1
        End If
    Next
    
    ' Optimisation sur lettre suivante
    OptimisationCol = col
    OptimisationLig = lig
    OptimisationOK = True

End Sub

Private Sub ComputeZoneBas(StrPosition As String, StrMaPioche As String, IntOrientation As Integer)
        
    Dim lig As Integer
    Dim col As Integer
    Dim i As Integer
    Dim IntStop As Integer
    Dim StrTypeCell As String
    Dim TabPossibilites(15) As String
    Dim T0 As Long
    Dim ParcoursTab As Integer
    Dim StrTemp As String
    Dim NbMots As Long
    Dim LgMini As Integer
    Dim LgMaxi As Integer
    Dim k As Long
    Dim StrMot As String
    Dim l As Integer
    Dim BlnOk As Boolean
    Dim StrEntete As String
    Dim StrSuffixe As String
    Dim Debut As Integer
    Dim Fin As Integer
    Dim MsTemp As ClsMotScribble
    Dim StrVraiMot As String
        
        
    lig = GetLigne(StrPosition)
    col = GetColonne(StrPosition)
        
    ' Déterminer s'il y a une zone Haute (possibilité d'ajouter une lettre au début d'un mot vertical )
    If lig <> 14 Then
        If Cell(col, lig + 1) = "" Then
                        
            '////////////////////////////////////////////////////////////////////////////
            
            Call GetTabPossibiliteFrom(col, lig + 1, Len(StrMaPioche), 0, TabPossibilites())
            
            For i = 0 To 14
                If TabPossibilites(i) <> "" Then
                    Debut = i
                    Exit For
                End If
            Next
            
            For i = Debut To 14
                If TabPossibilites(i) = "" Or i = 14 Then
                    Fin = i
                    Exit For
                End If
            Next
                        
            For ParcoursTab = Debut To col
                StrTemp = ""
                
                IntStop = Len(StrMaPioche)
                For i = ParcoursTab To Fin
                    If TabPossibilites(i) = "F" Then
                        StrTemp = StrTemp + Cell(i, lig + 1)
                    Else
                        StrTemp = StrTemp + "*"
                        IntStop = IntStop - 1
                        If IntStop = 0 Then Exit For
                    End If
                Next
                LgMini = col - ParcoursTab + 1
                LgMaxi = Fin - ParcoursTab + 2
                                
                NbMots = Ana.AnaGrammSelonMasque(StrTemp, StrMaPioche, LgMini)
                
                For k = 1 To NbMots
                    DoEvents
                    StrVraiMot = Ana.GetMot(k)
                    StrMot = SansJock(StrVraiMot)
                    
                    If Len(StrMot) <= LgMaxi Then
                        
                        BlnOk = True
                        '/////////////////////////////////////////////////////////////////////
                        'Optimisation
                        l = 1
                        Do
                            If TabPossibilites(ParcoursTab + l - 1) <> "L" Then Exit Do
                            l = l + 1
                            'if l>len(strmot) then msgbox "Impossible ???"
                        Loop While 1
                        '/////////////////////////////////////////////////////////////////////
                        
                        For l = l To Len(StrMot)
                            ' B
                            If InStr(TabPossibilites(ParcoursTab + l - 1), "B") Then
                                Cell(ParcoursTab + l - 1, lig + 1) = Mid(StrMot, l, 1)
                                ' le mot commence en lig-n
                                For i = lig To 0 Step -1
                                    If Cell(ParcoursTab + l - 1, i) = "" Then Exit For
                                Next
                                
                                If Ana.IsMotexiste(GetMot(Format(ParcoursTab + l) & Chr(i + 65 + 1))) = False Then
                                    Cell(ParcoursTab + l - 1, lig + 1) = ""
                                    BlnOk = False
                                    Exit For
                                End If
                                Cell(ParcoursTab + l - 1, lig + 1) = ""
                            End If
                            ' H
                            If InStr(TabPossibilites(ParcoursTab + l - 1), "H") Then
                                Cell(ParcoursTab + l - 1, lig + 1) = Mid(StrMot, l, 1)
                                For i = lig To 0 Step -1
                                    If Cell(ParcoursTab + l - 1, i) = "" Then Exit For
                                Next
                                
                                If Ana.IsMotexiste(GetMot(Format(ParcoursTab + l) & Chr(i + 65 + 1))) = False Then
                                    Cell(ParcoursTab + l - 1, lig + 1) = ""
                                    BlnOk = False
                                    Exit For
                                End If
                                Cell(ParcoursTab + l - 1, lig + 1) = ""
                            End If
                        Next
                        
                        ' Un mot est correct, mais verifier s'il y a des lettres fixes avant le mot ou après
                        If BlnOk Then
                            ' En-tete
                            StrEntete = ""
                            If ParcoursTab <> 0 Then
                                For i = ParcoursTab - 1 To 0 Step -1
                                    If TabPossibilites(i) <> "F" Then Exit For
                                    StrEntete = Cell(i, lig + 1) & StrEntete
                                Next
                            End If
                            
                            ' Suffixe
                            StrSuffixe = ""
                            For i = ParcoursTab + Len(StrMot) To 14
                                If TabPossibilites(i) <> "F" Then Exit For
                                StrSuffixe = StrSuffixe + Cell(i, lig + 1)
                            Next
                        
                            If Ana.IsMotexiste(StrEntete & StrMot & StrSuffixe) = False Then BlnOk = False
                        End If
                                                
                        If BlnOk Then
                        
                            Dim StrToAff As String
                            Dim points As Integer
                            Dim StrP As String
                            
                            StrToAff = ""
                            If StrEntete <> "" Then StrToAff = StrEntete
                            StrToAff = StrToAff & StrMot
                            If StrSuffixe <> "" Then StrToAff = StrToAff & StrSuffixe
                            StrToAff = LCase(StrToAff & AjouteJocker(StrVraiMot))
                            StrP = Chr(65 + lig + 1) & Format(ParcoursTab + 1)
                            Set TabSolution(PtrSolution) = ComptePointsMotHorizontal(StrToAff, StrP, IntOrientation)
                            PtrSolution = PtrSolution + 1
                        End If
                    End If
                Next
                    
                ' Optimisation
                Do
                    If TabPossibilites(ParcoursTab) = "F" Then
                        ParcoursTab = ParcoursTab + 1
                        If ParcoursTab > col Then Exit For
                    Else
                        Exit Do
                    End If
                Loop While 1
        
            Next
        End If
    End If

End Sub

Private Sub GetTabPossibiliteFrom(IntCol As Integer, IntLig As Integer, IntLgPicoche As Integer, DeltaCol, TabPossibilites() As String)

    Dim IntStop As Integer
    Dim i As Integer
    Dim StrTypeCell As String
    
    IntStop = IntLgPicoche
    
    For i = 0 To 14
        TabPossibilites(i) = ""
    Next
    
    For i = IntCol - 1 + DeltaCol To 0 Step -1
        StrTypeCell = ""
        If Cell(i, IntLig) = "" Then
            IntStop = IntStop - 1
            If IntStop = 0 Then Exit For
            ' H (Haut)
            If IntLig >= 1 Then
                If Cell(i, IntLig - 1) <> "" Then StrTypeCell = StrTypeCell + "H"
            End If
            ' B (Bas)
            If IntLig < 14 Then
                If Cell(i, IntLig + 1) <> "" Then StrTypeCell = StrTypeCell + "B"
            End If
            ' L (Libre)
            If StrTypeCell = "" Then StrTypeCell = "L"
        Else
            ' F (Fixe)
            StrTypeCell = "F"
        End If
        
        TabPossibilites(i) = StrTypeCell
    Next
                
    IntStop = IntLgPicoche + 1 - DeltaCol
    For i = IntCol + DeltaCol To 14
        StrTypeCell = ""
        If Cell(i, IntLig) = "" Then
            IntStop = IntStop - 1
            If IntStop = 0 Then Exit For
            ' H
            If IntLig >= 1 Then
                If Cell(i, IntLig - 1) <> "" Then StrTypeCell = StrTypeCell + "H"
            End If
            ' B
            If IntLig < 14 Then
                If Cell(i, IntLig + 1) <> "" Then StrTypeCell = StrTypeCell + "B"
            End If
            ' L (Libre)
            If StrTypeCell = "" Then StrTypeCell = "L"
        Else
            ' F (Fixe)
            StrTypeCell = "F"
        End If
        
        TabPossibilites(i) = StrTypeCell
    Next

End Sub

Private Function SansJock(StrMot As String) As String

    Dim p As Integer
    
    p = InStr(StrMot, " ")
    If p > 0 Then
        SansJock = Left(StrMot, p - 1)
    Else
        SansJock = StrMot
    End If
    
End Function

Private Function AjouteJocker(StrMot As String) As String

    Dim p As Integer
    
    p = InStr(StrMot, " ")
    If p > 0 Then
        AjouteJocker = Right(StrMot, Len(StrMot) - p + 1)
    Else
        AjouteJocker = ""
    End If
    
End Function

Private Function VerifierNecessiteTousJocker(ByRef StrMot As String, StrLettres As String) As Integer

    Dim i As Integer
    Dim j As Integer
    Dim k As Integer
    Dim StrTemp As String
    Dim StrJock As String
    Dim Car As String
    Dim p As Integer
    Dim Compte As Integer
    Dim Compte2 As Integer
    Dim NewJocks As String
        
    
    NewJocks = ""
    i = InStr(StrMot, " ")
    StrJock = Right(StrMot, Len(StrMot) - i - 1)
    StrJock = Left(StrJock, Len(StrJock) - 1)
    StrTemp = Left(StrMot, i - 1)
    
    For i = 1 To Len(StrJock)
        Car = Mid(StrJock, i, 1)
        
        Compte = 0
        For j = 1 To Len(StrLettres)
            If Mid(StrLettres, j, 1) = Car Then Compte = Compte + 1
        Next
        
        If Compte <> 0 Then
            ' Combien en ai je besoin dans Strmot ?
            For k = 1 To Len(StrTemp)
                If Mid(StrTemp, k, 1) = Car Then Compte2 = Compte2 + 1
            Next
            ' On peut substituer poour garder un '?'
            If Compte2 > Compte Then
                NewJocks = NewJocks + Car
            End If
        Else
            NewJocks = NewJocks + Car
        End If
    Next
        
    StrMot = StrTemp & " (" & NewJocks & ")"
    VerifierNecessiteTousJocker = Len(StrTemp)
    
End Function

Private Function PlacerPremierMot(Lettres As String) As Long

    Dim NbMots As Long
    Dim i As Long
    Dim j As Integer
    Dim StrMot As String
    Dim LgMot As Integer
    Dim StrPosition As String
        
    NbMots = Ana.AnaGramme(Lettres)
    
    For i = 1 To NbMots
        DoEvents
        StrMot = LCase(Ana.GetMot(i))
        
        j = InStr(StrMot, " ")
        If j Then
            LgMot = VerifierNecessiteTousJocker(StrMot, Lettres)
        Else
            LgMot = Len(StrMot)
        End If
        
        For j = 9 - LgMot To 8
            StrPosition = "H" & Format(j)
            Set TabSolution(PtrSolution) = ComptePointsMotHorizontal(StrMot, StrPosition, Int(Rnd() * 2))
            PtrSolution = PtrSolution + 1
        Next
    Next
    
    PlacerPremierMot = PtrSolution - 1

End Function

Public Function TrouveSolution(Lettres As String) As Long

    Dim Position As String
    Dim T0 As Long
    
    PtrSolution = 1
    ReDim TabSolution(100000)
    mSolutionTrouvee = False
    mIsSearching = True
    mAbort = False
       
    ' Si c'est le premier mot, trouver l'anagramme qui rapporte le plus de points
    If mPremierMot = True Then
        ' Pour les autres mots....
        TrouveSolution = PlacerPremierMot(Lettres)
        mSolutionTrouvee = True
        mIsSearching = False
        Exit Function
    End If
    
    ' Pour les autres mots.
    Position = ChercheLettre(True)
    Do
        Position = ChercheLettre()
        If Position = "" Then Exit Do
    
        If mAbort = False Then Call ComputeZoneHaut(Position, Lettres, 0)
        If mAbort = False Then Call ComputeZoneCentrale(Position, Lettres, 0)
        If mAbort = False Then Call ComputeZoneBas(Position, Lettres, 0)
        If mAbort Then Exit Do
        
    Loop While 1
    
    ' Inverser la grille
    SwapGrille
    
    Position = ChercheLettre(True)
    Do
        Position = ChercheLettre()
        If Position = "" Then Exit Do
    
        If mAbort = False Then Call ComputeZoneHaut(Position, Lettres, 1)
        If mAbort = False Then Call ComputeZoneCentrale(Position, Lettres, 1)
        If mAbort = False Then Call ComputeZoneBas(Position, Lettres, 1)
        
        If mAbort Then Exit Do
    Loop While 1
    
    ' Repasser la grille dans le sens normal
    SwapGrille
        
    TrouveSolution = PtrSolution - 1
    mSolutionTrouvee = True
    mIsSearching = False
        
End Function

Public Function GetSolution(IntNiveau As Integer) As ClsMotScribble

    ' Niveau de 1 à 5
    
    Dim i As Long
    Dim Max As Integer
    Dim PtrMax As Long
    Dim PtsCherches As Long
    Dim IntEcart As Integer
    Dim PtrSol As Long
    
    Max = 0
    For i = 1 To PtrSolution - 1
        If TabSolution(i).Pts > Max Then
            Max = TabSolution(i).Pts
            PtrMax = i
        End If
    Next
    
    If IntNiveau = 5 Then
        Set GetSolution = TabSolution(PtrMax)
        Exit Function
    End If
    
    'FrmMain.LstResultat.AddItem "Max : " & Format(Max)
    PtsCherches = (Max - 2) * IntNiveau / 5
    'FrmMain.LstResultat.AddItem "Points cherches : " & Format(PtsCherches)
    
    IntEcart = 9999
    For i = 1 To PtrSolution - 1
        If Abs(TabSolution(i).Pts - PtsCherches) < IntEcart Then
            IntEcart = Abs(TabSolution(i).Pts - PtsCherches)
            PtrSol = i
        End If
    Next
    
    Set GetSolution = TabSolution(PtrSol)

End Function

Property Get IsSolutionTrouvee() As Boolean

    IsSolutionTrouvee = mSolutionTrouvee
    
End Property

Property Let Abort(BlnState As Boolean)

    mAbort = BlnState
    
End Property

Property Get IsSearching() As Boolean
    
    IsSearching = mIsSearching

End Property
