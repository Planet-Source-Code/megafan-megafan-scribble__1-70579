VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ClsMots"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private TabResultat() As String
Private p As ClsPile
Private StrJocker As String
Private mDico As ClsDictionnaire
Private PtrResultat As Long

Public Function IsMotexiste(StrMot As String, Optional PtrPere As Long) As Boolean

    Dim PtrFils As Long
    Dim PtrPereBackup As Long
    Dim Index As Integer
    Dim lg As Integer
    Dim i As Integer
    
    StrMot = UCase(StrMot)
    If PtrPere = 0 Then
        Index = Asc(Left(StrMot, 1)) - 64
        PtrPere = Index
    End If
    lg = Len(StrMot)
    
    For i = 2 To lg
        Index = Asc(Mid(StrMot, i, 1)) - 64
        PtrFils = mDico.GetNoeud(PtrPere, Index)
        If PtrFils = 0 Then Exit Function
        PtrPereBackup = PtrPere
        PtrPere = PtrFils
    Next
    
    If mDico.MotExisteByNoeud(PtrPereBackup, Index) Then IsMotexiste = True

End Function

Public Function IsMotCommencePar(StrMot As String, Optional PtrPere As Long) As Long

    Dim PtrFils As Long
    Dim PtrPereBackup As Long
    Dim Index As Integer
    Dim lg As Integer
    Dim i As Integer
    
    If PtrPere = 0 Then
        Index = Asc(Left(StrMot, 1)) - 64
        PtrPere = Index
    End If
    lg = Len(StrMot)
    
    For i = 2 To lg
        Index = Asc(Mid(StrMot, i, 1)) - 64
        PtrFils = mDico.GetNoeud(PtrPere, Index)
        If PtrFils = 0 Then Exit Function
        PtrPereBackup = PtrPere
        PtrPere = PtrFils
    Next
    
    IsMotCommencePar = PtrPere

End Function

Public Function GetMot(LngIndex As Long) As String

    GetMot = TabResultat(LngIndex)

End Function

Public Function TotalMots() As Long

    TotalMots = UBound(TabResultat)

End Function

Public Function AnaGramme(StrMot As String, Optional MotComplet As String, Optional PtrArbre As Long) As Long

    Dim i As Long
    Dim p As Integer
    ReDim TabResultat(0)
    ReDim TabResultat(MAXMOTS)
    PtrResultat = 1
    StrJocker = ""
    
    '//////// Mettre les ? à la fin de StrMot/////////////////////////////
    i = 0
    Do
        p = InStr(StrMot, "?")
        If p Then
            StrMot = Left(StrMot, p - 1) & Right(StrMot, Len(StrMot) - p)
            i = i + 1
        Else
            Exit Do
        End If
    Loop While 1
    StrMot = StrMot & String(i, "?")
    '////////////////////////////////////////////////////////////////////
    
    AnaGrammeP UCase(StrMot)
   
    For i = 1 To UBound(TabResultat)
        If TabResultat(i) = "" Then Exit For
    Next
    ReDim Preserve TabResultat(i - 1)
    
    Me.ClasserMotsParLongueur
    
    AnaGramme = i - 1

End Function

Private Sub AnaGrammeP(StrMot As String, Optional MotComplet As String, Optional PtrArbre As Long)
        
    Dim i As Integer
    Dim IndexFils As Integer
    Dim TempMot As String
    Dim Lettre As String
    Dim LettresUtilisees  As String
    Dim k As Integer
    
    
    For i = 1 To Len(StrMot)
        Lettre = Mid(StrMot, i, 1)
        If Lettre <> "?" Then
            If InStr(LettresUtilisees, Lettre) = 0 Then
                IndexFils = Asc(Left(Lettre, 1)) - 64
                MotComplet = MotComplet & Lettre
                LettresUtilisees = LettresUtilisees & Lettre
                
                TempMot = StrMot
                TempMot = Left(TempMot, i - 1) & Right(TempMot, Len(TempMot) - i)
                
                If mDico.MotExisteByNoeud(PtrArbre, IndexFils) Then
                    AjouteListe MotComplet, StrJocker
                End If
            
                If TempMot <> "" Then
                    If mDico.GetNoeud(PtrArbre, IndexFils) <> 0 Then
                        AnaGrammeP TempMot, MotComplet, mDico.GetNoeud(PtrArbre, IndexFils)
                    End If
                End If
                MotComplet = Left(MotComplet, Len(MotComplet) - 1)
            End If
        Else
            For k = 65 To 90 ' A to Z
                Lettre = Chr(k)
                If InStr(LettresUtilisees, Lettre) = 0 Then
                    IndexFils = Asc(Left(Lettre, 1)) - 64
                    MotComplet = MotComplet & Lettre
                    StrJocker = StrJocker & Lettre
                    LettresUtilisees = LettresUtilisees & Lettre
                    
                    TempMot = StrMot
                    TempMot = Left(TempMot, i - 1) & Right(TempMot, Len(TempMot) - i)
                
                    If mDico.MotExisteByNoeud(PtrArbre, IndexFils) Then
                        AjouteListe MotComplet, StrJocker
                    End If
                
                    If TempMot <> "" Then
                        If mDico.GetNoeud(PtrArbre, IndexFils) <> 0 Then
                            AnaGrammeP TempMot, MotComplet, mDico.GetNoeud(PtrArbre, IndexFils)
                        End If
                    End If
                    MotComplet = Left(MotComplet, Len(MotComplet) - 1)
                    StrJocker = Left(StrJocker, Len(StrJocker) - 1)
                End If
            Next
        End If
    Next
    
End Sub

Public Function AnaGrammSelonMasque(StrMasque As String, StrMot As String, IntLgMini As Integer) As Long
    
    ReDim TabResultat(0)
    ReDim TabResultat(MAXMOTS)
    PtrResultat = 1
    StrJocker = ""
    Dim i As Long
    Dim p As Integer
    
    '//////// Mettre les ? à la fin de StrMot/////////////////////////////
    i = 0
    Do
        p = InStr(StrMot, "?")
        If p Then
            StrMot = Left(StrMot, p - 1) & Right(StrMot, Len(StrMot) - p)
            i = i + 1
        Else
            Exit Do
        End If
    Loop While 1
    StrMot = StrMot & String(i, "?")
    '////////////////////////////////////////////////////////////////////
    
    
    AnaGrammSelonMasqueP UCase(StrMasque), UCase(StrMot), IntLgMini

    For i = 1 To UBound(TabResultat)
        If TabResultat(i) = "" Then Exit For
    Next
    ReDim Preserve TabResultat(i - 1)
        
    AnaGrammSelonMasque = i - 1

End Function


Private Sub AnaGrammSelonMasqueP(StrMasque As String, StrMot As String, IntLgMini As Integer, Optional MotComplet As String, Optional PtrArbre As Long, Optional PtrMasque As Integer)
        
    Dim i As Integer
    Dim IndexFils As Integer
    Dim TempMot As String
    Dim Lettre As String
    Dim LettresUtilisees  As String
    Dim k As Integer
    Dim SavePtr As Long
    Dim StrTemp As String
    
        
    For i = 1 To Len(StrMot)
        Lettre = Mid(StrMot, i, 1)
        If Lettre <> "?" Then
            If InStr(LettresUtilisees, Lettre) = 0 Then
                If PtrMasque < Len(StrMasque) Then
                    PtrMasque = PtrMasque + 1
                    If Mid(StrMasque, PtrMasque, 1) <> "*" Then
                        Do
                            Lettre = Mid(StrMasque, PtrMasque, 1)
                            IndexFils = Asc(Left(Lettre, 1)) - 64
                            If mDico.GetNoeud(PtrArbre, IndexFils) <> 0 Then
                                MotComplet = MotComplet + Lettre
                                PtrMasque = PtrMasque + 1
                                If mDico.MotExisteByNoeud(PtrArbre, IndexFils) Then
                                    If Len(MotComplet) >= IntLgMini Then
                                        AjouteListe MotComplet, StrJocker
                                    End If
                                End If
                                
                                p.Push PtrArbre
                                PtrArbre = mDico.GetNoeud(PtrArbre, IndexFils)
                                
                                If Len(MotComplet) < Len(StrMasque) Then
                                    Lettre = Mid(StrMot, i, 1)
                                Else
                                    PtrMasque = PtrMasque - 1
                                    Lettre = ""
                                    Exit Do
                                End If
                            Else
                                PtrMasque = PtrMasque - 1
                                Lettre = ""
                                Exit Do
                            End If
                            
                        Loop While Mid(StrMasque, PtrMasque, 1) <> "*"
                    End If

                    If Lettre <> "" Then
                        MotComplet = MotComplet & Lettre
                        IndexFils = Asc(Left(Lettre, 1)) - 64
                        LettresUtilisees = LettresUtilisees & Lettre
                        
                        TempMot = StrMot
                        TempMot = Left(TempMot, i - 1) & Right(TempMot, Len(TempMot) - i)
                        
                        If mDico.MotExisteByNoeud(PtrArbre, IndexFils) Then
                            If Len(MotComplet) >= IntLgMini Then
                                AjouteListe MotComplet, StrJocker
                            End If
                        End If
                    
                        'If TempMot <> "" Then
                            If PtrMasque < Len(StrMasque) Then
                                If mDico.GetNoeud(PtrArbre, IndexFils) <> 0 Then
                                    AnaGrammSelonMasqueP StrMasque, TempMot, IntLgMini, MotComplet, mDico.GetNoeud(PtrArbre, IndexFils), PtrMasque
                                End If
                            End If
                        'End If
                        MotComplet = Left(MotComplet, Len(MotComplet) - 1)
                        PtrMasque = PtrMasque - 1
                    End If
                End If
            End If
        Else
            For k = 65 To 90 ' A to Z
                DoEvents '<---------------------------------------- Indispensable
                Lettre = Chr(k)
                If InStr(LettresUtilisees, Lettre) = 0 Then
                    If PtrMasque < Len(StrMasque) Then
                        PtrMasque = PtrMasque + 1
                        If Mid(StrMasque, PtrMasque, 1) <> "*" Then
                            Do
                                'DoEvents
                                Lettre = Mid(StrMasque, PtrMasque, 1)
                                IndexFils = Asc(Left(Lettre, 1)) - 64
                                If mDico.GetNoeud(PtrArbre, IndexFils) <> 0 Then
                                    MotComplet = MotComplet + Lettre
                                    PtrMasque = PtrMasque + 1
                                    If mDico.MotExisteByNoeud(PtrArbre, IndexFils) Then
                                        If Len(MotComplet) >= IntLgMini Then
                                            AjouteListe MotComplet, StrJocker
                                        End If
                                    End If
                                    
                                    p.Push PtrArbre
                                    PtrArbre = mDico.GetNoeud(PtrArbre, IndexFils)
                                    
                                    If Len(MotComplet) < Len(StrMasque) Then
                                        Lettre = Mid(StrMot, i, 1)
                                        Lettre = Chr(k)
                                    Else
                                        PtrMasque = PtrMasque - 1
                                        Lettre = ""
                                        Exit Do
                                    End If
                                Else
                                    PtrMasque = PtrMasque - 1
                                    Lettre = ""
                                    Exit Do
                                End If
                                
                            Loop While Mid(StrMasque, PtrMasque, 1) <> "*"
                        End If
                        

                        If Lettre <> "" Then
                        
                            StrJocker = StrJocker & Lettre
                            MotComplet = MotComplet & Lettre
                            IndexFils = Asc(Left(Lettre, 1)) - 64
                            LettresUtilisees = LettresUtilisees & Lettre
                            
                            TempMot = StrMot
                            TempMot = Left(TempMot, i - 1) & Right(TempMot, Len(TempMot) - i)
                            
                            If mDico.MotExisteByNoeud(PtrArbre, IndexFils) Then
                                If Len(MotComplet) >= IntLgMini Then
                                    AjouteListe MotComplet, StrJocker
                                End If
                            End If
                        
                            'If TempMot <> "" Then
                                If PtrMasque < Len(StrMasque) Then
                                    If mDico.GetNoeud(PtrArbre, IndexFils) <> 0 Then
                                        AnaGrammSelonMasqueP StrMasque, TempMot, IntLgMini, MotComplet, mDico.GetNoeud(PtrArbre, IndexFils), PtrMasque
                                    End If
                                End If
                            'End If
                            MotComplet = Left(MotComplet, Len(MotComplet) - 1)
                            PtrMasque = PtrMasque - 1
                            StrJocker = Left(StrJocker, Len(StrJocker) - 1)
                        End If
                    End If
                End If
                'StrJocker = Left(StrJocker, Len(StrJocker) - 2)
            Next
        End If
    Next
    
    If MotComplet <> "" Then
        If Mid(StrMasque, Len(MotComplet), 1) <> "*" Then
            Do
                MotComplet = Left(MotComplet, Len(MotComplet) - 1)
                PtrMasque = PtrMasque - 1
                p.Pop PtrArbre
                If MotComplet = "" Then Exit Do
            Loop While Mid(StrMasque, Len(MotComplet), 1) <> "*"
        End If
        
        If StrMot = "" Then
            If Len(MotComplet) < Len(StrMasque) Then
                StrTemp = Right(StrMasque, Len(StrMasque) - Len(MotComplet))
                If InStr(StrTemp, "*") = 0 Then
                    If IsMotexiste(MotComplet & StrTemp) Then
                        If Len(MotComplet) >= IntLgMini Then
                            AjouteListe MotComplet, StrJocker
                        End If
                    End If
                End If
            End If
        End If
    End If
    
End Sub

Private Sub AjouteListe(StrMot As String, Optional StrJocker As String)
        
    If StrJocker <> "" Then
        TabResultat(PtrResultat) = StrMot & " (" & StrJocker & ")"
    Else
        TabResultat(PtrResultat) = StrMot
    End If
    PtrResultat = PtrResultat + 1
    
End Sub

Public Sub ClasserMotParOrdreAlphabetique()

    Dim i As Long
    Dim j As Long
    Dim StrTemp As String
    
    If UBound(TabResultat) = 0 Then Exit Sub
    If UBound(TabResultat) > 10000 Then
        MsgBox "Il y a " & CStr(UBound(TabResultat)) & " solutions, veuillez préciser votre recherche !", vbExclamation
        Exit Sub
    End If

    For i = 1 To UBound(TabResultat)
        For j = i + 1 To UBound(TabResultat)
            If j <> i Then
                If TabResultat(j) < TabResultat(i) Then
                    StrTemp = TabResultat(j)
                    TabResultat(j) = TabResultat(i)
                    TabResultat(i) = StrTemp
                End If
            End If
        Next
    Next

End Sub

Public Sub ClasserMotsParLongueur()

    Dim i As Long
    Dim j As Long
    Dim StrTemp As String
    Dim LgMot As Integer
    Dim Debut As Long
    
    If UBound(TabResultat) = 0 Then Exit Sub
    If UBound(TabResultat) > 10000 Then
        MsgBox "Il y a " & CStr(UBound(TabResultat)) & " solutions, veuillez préciser votre recherche !", vbExclamation
        Exit Sub
    End If
    
    For i = 1 To UBound(TabResultat)
        For j = 1 To UBound(TabResultat)
            If Len(TabResultat(j)) < Len(TabResultat(i)) Then
                StrTemp = TabResultat(j)
                TabResultat(j) = TabResultat(i)
                TabResultat(i) = StrTemp
            End If
        Next
    Next
End Sub

Private Sub Class_Initialize()
    
    Set p = New ClsPile
    p.InitPile (50)
      
End Sub

Property Let Dictionnaire(Dico As ClsDictionnaire)

   Set mDico = Dico
   
End Property


