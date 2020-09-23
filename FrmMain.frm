VERSION 5.00
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form FrmMain 
   AutoRedraw      =   -1  'True
   BackColor       =   &H00A36C38&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Megafan Scribble"
   ClientHeight    =   12165
   ClientLeft      =   4530
   ClientTop       =   3945
   ClientWidth     =   15165
   BeginProperty Font 
      Name            =   "Arial"
      Size            =   9.75
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   ForeColor       =   &H00000000&
   Icon            =   "FrmMain.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   811
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   1011
   Begin VB.Timer TimTourSuivant 
      Enabled         =   0   'False
      Interval        =   500
      Left            =   600
      Top             =   9360
   End
   Begin VB.Timer TimChrono 
      Enabled         =   0   'False
      Interval        =   300
      Left            =   600
      Top             =   8880
   End
   Begin VB.Frame FrameChronometre 
      BackColor       =   &H00A36C38&
      Caption         =   "Chronomètre"
      Height          =   855
      Left            =   9000
      TabIndex        =   24
      Top             =   6480
      Width           =   4395
      Begin VB.CommandButton CmdFinTour 
         Caption         =   "Fin du tour"
         Enabled         =   0   'False
         Height          =   315
         Left            =   2280
         TabIndex        =   25
         Top             =   350
         Width           =   1815
      End
      Begin VB.Label LblChrono 
         Alignment       =   2  'Center
         BackColor       =   &H0080FFFF&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "00:00"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   14.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   375
         Left            =   360
         TabIndex        =   26
         Top             =   300
         Width           =   1455
      End
   End
   Begin VB.CommandButton CmdChanger 
      Caption         =   "Changer"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   675
      Left            =   12480
      TabIndex        =   23
      Top             =   10680
      Width           =   1455
   End
   Begin VB.Frame FrameScores 
      BackColor       =   &H00A36C38&
      Caption         =   "Scores"
      Height          =   1335
      Left            =   9000
      TabIndex        =   18
      Top             =   5040
      Width           =   4395
      Begin VB.TextBox TxtPts 
         Alignment       =   2  'Center
         BackColor       =   &H0080FFFF&
         Height          =   315
         Left            =   2880
         Locked          =   -1  'True
         MousePointer    =   1  'Arrow
         TabIndex        =   20
         Top             =   840
         Width           =   735
      End
      Begin VB.TextBox TxtPtsJoueur 
         Alignment       =   2  'Center
         BackColor       =   &H0080FFFF&
         Height          =   315
         Left            =   2880
         Locked          =   -1  'True
         MousePointer    =   1  'Arrow
         TabIndex        =   19
         Top             =   360
         Width           =   735
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Points Joueur"
         Height          =   255
         Left            =   1080
         TabIndex        =   22
         Top             =   360
         Width           =   1695
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "Points Ordinateur"
         Height          =   255
         Left            =   1080
         TabIndex        =   21
         Top             =   840
         Width           =   1695
      End
   End
   Begin VB.Frame FrameHistorique 
      BackColor       =   &H00A36C38&
      Caption         =   "Historique partie"
      Height          =   4575
      Left            =   9000
      TabIndex        =   16
      Top             =   360
      Width           =   4395
      Begin MSHierarchicalFlexGridLib.MSHFlexGrid GridHistorique 
         Height          =   4050
         Left            =   120
         TabIndex        =   17
         Top             =   360
         Width           =   4130
         _ExtentX        =   7276
         _ExtentY        =   7144
         _Version        =   393216
         BackColor       =   12648447
         FixedCols       =   0
         BackColorBkg    =   8454143
         FocusRect       =   0
         HighLight       =   0
         ScrollBars      =   2
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         _NumberOfBands  =   1
         _Band(0).Cols   =   2
      End
   End
   Begin VB.Frame FrameOrthographe 
      BackColor       =   &H00A36C38&
      Caption         =   "Vérification dictionnaire"
      Height          =   1095
      Left            =   9120
      TabIndex        =   14
      Top             =   8040
      Width           =   4395
      Begin VB.PictureBox PicOK 
         AutoSize        =   -1  'True
         BorderStyle     =   0  'None
         Height          =   315
         Left            =   2880
         Picture         =   "FrmMain.frx":5F32
         ScaleHeight     =   315
         ScaleWidth      =   315
         TabIndex        =   28
         Top             =   360
         Width           =   315
      End
      Begin VB.PictureBox PicKO 
         AutoSize        =   -1  'True
         BorderStyle     =   0  'None
         Height          =   315
         Left            =   2880
         Picture         =   "FrmMain.frx":64B4
         ScaleHeight     =   315
         ScaleWidth      =   315
         TabIndex        =   27
         Top             =   360
         Width           =   315
      End
      Begin VB.TextBox TxtMotAVerifier 
         BackColor       =   &H0080FFFF&
         Height          =   315
         Left            =   720
         TabIndex        =   15
         Top             =   360
         Width           =   1935
      End
   End
   Begin VB.PictureBox PicLettreJoueur 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   20.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   495
      Index           =   3
      Left            =   4920
      ScaleHeight     =   33
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   33
      TabIndex        =   13
      Top             =   10320
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.CommandButton CmdMelanger 
      Caption         =   "Mélanger"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   10920
      TabIndex        =   12
      Top             =   11040
      Width           =   1455
   End
   Begin VB.CommandButton CmdRAZJetons 
      Caption         =   "RAZ Jetons"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   10920
      TabIndex        =   11
      Top             =   10680
      Width           =   1455
   End
   Begin VB.CommandButton CmdPasser 
      Caption         =   "Passer"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   9360
      TabIndex        =   10
      Top             =   11040
      Width           =   1455
   End
   Begin MSComctlLib.StatusBar StatusBar 
      Align           =   2  'Align Bottom
      Height          =   375
      Left            =   0
      TabIndex        =   9
      Top             =   11790
      Width           =   15165
      _ExtentX        =   26749
      _ExtentY        =   661
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   2
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Object.Width           =   24342
            MinWidth        =   24342
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            Object.Width           =   2249
            MinWidth        =   2249
         EndProperty
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.CommandButton CmdJouer 
      Caption         =   "Jouer"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   9360
      TabIndex        =   8
      Top             =   10680
      Width           =   1455
   End
   Begin VB.Timer TimDrag 
      Enabled         =   0   'False
      Interval        =   1
      Left            =   600
      Top             =   8400
   End
   Begin VB.PictureBox PicLettreJoueur 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   20.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   495
      Index           =   6
      Left            =   6720
      ScaleHeight     =   33
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   33
      TabIndex        =   7
      Top             =   10320
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.PictureBox PicLettreJoueur 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   20.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   495
      Index           =   5
      Left            =   6120
      ScaleHeight     =   33
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   33
      TabIndex        =   6
      Top             =   10320
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.PictureBox PicLettreJoueur 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   20.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   495
      Index           =   4
      Left            =   5520
      ScaleHeight     =   33
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   33
      TabIndex        =   5
      Top             =   10320
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.PictureBox PicLettreJoueur 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   20.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   495
      Index           =   2
      Left            =   4320
      ScaleHeight     =   33
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   33
      TabIndex        =   4
      Top             =   10320
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.PictureBox PicLettreJoueur 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   20.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   495
      Index           =   1
      Left            =   3720
      ScaleHeight     =   33
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   33
      TabIndex        =   3
      Top             =   10320
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.PictureBox PicLettreJoueur 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   20.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   495
      Index           =   0
      Left            =   3120
      ScaleHeight     =   33
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   33
      TabIndex        =   2
      Top             =   10320
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.PictureBox PicReglette 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   15.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   615
      Left            =   480
      ScaleHeight     =   41
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   385
      TabIndex        =   1
      Top             =   6480
      Width           =   5775
   End
   Begin VB.PictureBox PicGrille 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   15.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   5535
      Left            =   480
      ScaleHeight     =   369
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   385
      TabIndex        =   0
      Top             =   480
      Width           =   5775
   End
   Begin VB.Menu MnuPartie 
      Caption         =   "Partie"
      Begin VB.Menu MnuEnregistrerPartie 
         Caption         =   "&Enregistrer une partie..."
         Visible         =   0   'False
      End
      Begin VB.Menu MnuChargerPartie 
         Caption         =   "&Charger une partie..."
         Visible         =   0   'False
      End
      Begin VB.Menu MnuRien4 
         Caption         =   "-"
         Visible         =   0   'False
      End
      Begin VB.Menu MnuNouvellePartie 
         Caption         =   "&Nouvelle partie ..."
      End
      Begin VB.Menu MnuArreterLaPartie 
         Caption         =   "&Arrêter la partie..."
         Visible         =   0   'False
      End
      Begin VB.Menu MnuRien1 
         Caption         =   "-"
      End
      Begin VB.Menu MnuQuitter 
         Caption         =   "&Quitter"
      End
   End
   Begin VB.Menu MnuOutils 
      Caption         =   "&Outils"
      Begin VB.Menu MnuAnagrammes 
         Caption         =   "&Anagrammes..."
      End
      Begin VB.Menu MnuSolution 
         Caption         =   "Solution..."
         Enabled         =   0   'False
      End
   End
   Begin VB.Menu MnuParamertes 
      Caption         =   "Paramètres"
      Begin VB.Menu MnuNiveauPC 
         Caption         =   "&Niveau difficulté PC..."
      End
      Begin VB.Menu Mnu1erTour 
         Caption         =   "&Définir le 1er tour..."
      End
      Begin VB.Menu MnuAssistance 
         Caption         =   "&Barre d'état..."
         Visible         =   0   'False
      End
      Begin VB.Menu MnuRien3 
         Caption         =   "-"
      End
      Begin VB.Menu MnuModeDeJeu 
         Caption         =   "&Mode de jeu..."
      End
      Begin VB.Menu MnuRien2 
         Caption         =   "-"
      End
      Begin VB.Menu MnuChangerDico 
         Caption         =   "&Changer de dictionnaire..."
      End
   End
   Begin VB.Menu MnuAide 
      Caption         =   "?"
      Begin VB.Menu MnuAPropos 
         Caption         =   "&A propos de..."
      End
   End
End
Attribute VB_Name = "FrmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'
'
'       TODO :
'
'       - Gestion de la barre d'état : MnuBarreEtat - FrmBarreEtat
'       - Gestion des points des images / dictionnaires (ex J8 ODS alors que J5 en TWL)
'       - Fichier Enregistrer/Ouvrir
'       - Mode x joueur (humain + PC ????)
'
'
Option Explicit

Private Declare Function timeGetTime Lib "winmm.dll" () As Long
Private TailleJetons As Integer
Private Ana As ClsMots
Private MsDuplicate As ClsMotScribble
Private Tour As EPremierTour
Private G As ClsGrille
Private GDuplicate As ClsGrille
Private NbJePasse As Integer

Private Type MvtLettre
    Dx As Integer
    Dy As Integer
    Idx As Integer
    OldX As Integer
    OldY As Integer
    BorderDiffX As Integer
    BorderDiffY As Integer
    PosBlocLettre As Integer
End Type

Private Type LP
    x As Integer
    y As Integer
    Car As String
End Type

Private Type TReglette
    Lettre As String
    IndexImg As Integer
End Type

Dim TabReglette(12) As TReglette
Dim LettreJoueur As MvtLettre
    
Private Type TMelange
    IndexBloc  As Integer
    Lettre As String
    IndexImg As Integer
    Left As Integer
End Type

Private LETTRESUC As String

Private Function CompterPoints() As ClsMotScribble

    Dim i As Integer
    Dim j As Integer
    Dim Ligne As Integer
    Dim Colonne As Integer
    Dim StrMot As String
    Dim StrTemp As String
    Dim LettresPlacees As Integer
    Dim TabLettresPlacees(7) As LP
    Dim BlnErr As Boolean
    Dim Orientation As Integer
    Dim GJoueur As ClsGrille
    
    Set CompterPoints = New ClsMotScribble
    Set GJoueur = New ClsGrille
    
    CmdJouer.Enabled = False
    
    Debug.Print
    Debug.Print
    Debug.Print
    Debug.Print
    Debug.Print
    
    LettresPlacees = 0
    For i = 0 To 6
        If PicLettreJoueur(i).Tag <> "" Then
            ' Cette image est dans la grille ?
            If PicLettreJoueur(i).Left + PicLettreJoueur(0).Width >= PicGrille.Left And PicLettreJoueur(i).Left < PicGrille.Left + PicGrille.Width Then
                If PicLettreJoueur(i).Top + PicLettreJoueur(0).Height >= PicGrille.Top And PicLettreJoueur(i).Top < PicGrille.Top + PicGrille.Height Then
                    
                    Ligne = (PicLettreJoueur(i).Top - PicGrille.Top) / ((PicGrille.Height - 2) / 15)
                    Colonne = (PicLettreJoueur(i).Left - PicGrille.Left) / ((PicGrille.Width - 2) / 15)
                    
                    TabLettresPlacees(LettresPlacees).x = Colonne
                    TabLettresPlacees(LettresPlacees).y = Ligne
                    TabLettresPlacees(LettresPlacees).Car = PicLettreJoueur(i).Tag
                    LettresPlacees = LettresPlacees + 1
                End If
            End If
        End If
    Next
    
    If LettresPlacees = 0 Then Exit Function
    
    Ligne = TabLettresPlacees(0).y
    Colonne = TabLettresPlacees(0).x
    Orientation = -1
    
    If LettresPlacees > 1 Then
        ' Plusieurs lettres placées
        Debug.Print Format(LettresPlacees) & " lettres placées."
        ' Déterminer l'orientation
        Orientation = -1
        If TabLettresPlacees(1).y = Ligne Then Orientation = 0
        If TabLettresPlacees(1).x = Colonne Then Orientation = 1
        If Orientation = 0 Then
            Debug.Print "Orientation horizontale"
        Else
            Debug.Print "Orientation verticale"
        End If
        If Orientation = -1 Then
            BlnErr = True
            Debug.Print "Sortie : Orientation indeterminée..."
        Else
            BlnErr = False
        End If
        
        For i = 2 To LettresPlacees - 1
            If Orientation = 0 Then
                If TabLettresPlacees(i).y <> Ligne Then BlnErr = True
            Else
                If TabLettresPlacees(i).x <> Colonne Then BlnErr = True
            End If
        Next
        
        If BlnErr Then
            Debug.Print "Sortie : X ou Y d'une lettre different."
            Exit Function
        End If
    Else
        ' 1 seule lettre placée.
        Debug.Print "1 lettres placée."
        ' pas au premier tour ?
        If G.IsPremierMot Then
            Debug.Print "Sortie : 1 lettres placée au premier tour !"
            Exit Function
        End If
        ' Déterminer l'orientation
        Orientation = -1
        If Colonne < 14 Then
            ' Vérifier une lettre à droite
            If G.Cell(Colonne + 1, Ligne) <> "" Then Orientation = 0
            Debug.Print "Orientation horizontale"
        End If
        
        If Colonne > 0 Then
            ' Vérifier une lettre à gauche
            If G.Cell(Colonne - 1, Ligne) <> "" Then Orientation = 0
            Debug.Print "Orientation horizontale"
        End If
        
        If Ligne < 14 Then
            ' Vérifier une lettre en dessous
            If G.Cell(Colonne, Ligne + 1) <> "" Then Orientation = 1
            Debug.Print "Orientation verticale"
        End If
        
        If Ligne > 0 Then
            ' Vérifier une lettre en haut
            If G.Cell(Colonne, Ligne - 1) <> "" Then Orientation = 1
            Debug.Print "Orientation verticale"
        End If
        
        If Orientation = -1 Then Exit Function
    
    End If
    
    ' Remplir la grille GJoueur
    For i = 0 To LettresPlacees - 1
        GJoueur.Cell(TabLettresPlacees(i).x, TabLettresPlacees(i).y) = TabLettresPlacees(i).Car
    Next
    
    ' Vérifier contact avec une lettre déja placée...
    If G.IsPremierMot Then
        ' vérifier qu'il y a une lettre placée en (7,7)
        If GJoueur.Cell(7, 7) = "" Then
            Debug.Print "Sortie : Pas de lettre en I8"
            Exit Function
        End If
    Else
        ' Vérifier un contact avec une lettre de la grille
        For i = 0 To LettresPlacees - 1
            ' contact à gauche
            If TabLettresPlacees(i).x > 0 Then
                If G.Cell(TabLettresPlacees(i).x - 1, TabLettresPlacees(i).y) <> "" Then Exit For
            End If
            ' contact à droite
            If TabLettresPlacees(i).x < 14 Then
                If G.Cell(TabLettresPlacees(i).x + 1, TabLettresPlacees(i).y) <> "" Then Exit For
            End If
            ' contact en haut
            If TabLettresPlacees(i).y > 0 Then
                If G.Cell(TabLettresPlacees(i).x, TabLettresPlacees(i).y - 1) <> "" Then Exit For
            End If
            ' contact en bas
            If TabLettresPlacees(i).y < 14 Then
                If G.Cell(TabLettresPlacees(i).x, TabLettresPlacees(i).y + 1) <> "" Then Exit For
            End If
            
        Next
        
        If i >= LettresPlacees Then
            Debug.Print "Sortie : Pas de contact avec les autres mots"
            Exit Function
        End If
    End If
     
    ' Si l'orientation est verticale, inverser la grille
    If Orientation = 1 Then
        GJoueur.SwapGrille
        G.SwapGrille
        i = Colonne
        Colonne = Ligne
        Ligne = i
        Debug.Print "Inversion de la grille."
    End If
    
    ' Déterminer le début du mot (à gauche)
    Do
        If GJoueur.Cell(Colonne, Ligne) <> "" Or G.Cell(Colonne, Ligne) <> "" Then
            Colonne = Colonne - 1
        Else
            Exit Do
        End If
    Loop While Colonne >= 0
    
    Colonne = Colonne + 1
    Debug.Print "Le mot commence en colonne " & Format(Colonne)
    
    ' Créer le mot à vérifier
    StrMot = ""
    For i = Colonne To 14
        If GJoueur.Cell(i, Ligne) = "" And G.Cell(i, Ligne) = "" Then Exit For
        If GJoueur.Cell(i, Ligne) <> "" Then
            StrMot = StrMot + GJoueur.Cell(i, Ligne)
            LettresPlacees = LettresPlacees - 1
        Else
            StrMot = StrMot + G.Cell(i, Ligne)
        End If
    Next
    
    ' Toutes les lettres placées sont alignées ?
    If LettresPlacees <> 0 Then
        ' Remettre la grille dans le sens horizontal si besoin
        If Orientation = 1 Then G.SwapGrille
        Debug.Print "Sortie : Certaines lettres n'ont pas ete utilisées dans le mot"
        Exit Function
    End If
            
    ' Le mot principal existe ?
    If Ana.IsMotexiste(UCase(StrMot)) = False Then
        ' Remettre la grille dans le sens horizontal si besoin
        If Orientation = 1 Then G.SwapGrille
        Debug.Print "Sortie : '" & StrMot & "' n'existe pas !"
        Exit Function
    End If
    Debug.Print "Mot à calculer :'" & StrMot & "'"
    
    ' Les mots verticaux existent ?
    BlnErr = False
    For i = Colonne To 14
        If GJoueur.Cell(i, Ligne) = "" And G.Cell(i, Ligne) = "" Then Exit For
        ' Ne vérifier que les mots crees
        If GJoueur.Cell(i, Ligne) <> "" Then
            ' transférer la lettre provisoirement
            G.Cell(i, Ligne) = Mid(StrMot, i - Colonne + 1, 1)
           
            j = Ligne
            Do While G.Cell(i, j) <> ""
                j = j - 1
                If j = -1 Then Exit Do
            Loop
            
            StrTemp = G.GetMot(Format(i + 1) & Chr(j + 66))
            G.Cell(i, Ligne) = ""
            
            If Len(StrTemp) > 1 Then
                Debug.Print "vérification du mot : '" & StrTemp & "'"
                If Ana.IsMotexiste(StrTemp) = False Then
                    BlnErr = True
                    Debug.Print "'" & StrTemp & "' n'existe pas !"
                    Exit For
                End If
            End If
            
        End If
    Next
    
    If BlnErr = True Then
        ' Remettre la grille dans le sens horizontal si besoin
        If Orientation = 1 Then G.SwapGrille
        Exit Function
    End If
    
    Debug.Print "Placé en : " & Chr(Ligne + 65) & Format(Colonne + 1)
    Set CompterPoints = G.ComptePointsMotHorizontal(StrMot, Chr(Ligne + 65) & Format(Colonne + 1), Orientation)
    
    ' Remettre la grille dans le sens horizontal si besoin
    If Orientation = 1 Then G.SwapGrille
    
    ' Le coup est possible
    CmdJouer.Enabled = True
    
End Function

Private Sub CmdChanger_Click()

    Dim TabIndex() As String
    Dim StrLettresAChanger As String
    Dim i As Integer
    Dim j As Integer
    
    ' remettre les lettres sur la reglette
    CmdRAZJetons_Click
    
    TimChrono.Enabled = False
    FrmChange.Show vbModal
    
    If FrmChange.Tag <> "" Then
        TabIndex() = Split(FrmChange.Tag, ",")
    Else
        ReDim TabIndex(0)
        If PARTIECHRONOMETREE = "oui" Then TimChrono.Enabled = True
        TabIndex(0) = -1
    End If
    
    Unload FrmChange
    If TabIndex(0) = -1 Then Exit Sub
        
    For i = 0 To UBound(TabIndex) - 1
        StrLettresAChanger = StrLettresAChanger & UCase(PicLettreJoueur(CInt(TabIndex(i))).Tag)
        PicLettreJoueur(CInt(TabIndex(i))).Tag = ""
        For j = 0 To 12
            If TabReglette(j).IndexImg = CInt(TabIndex(i)) Then
                TabReglette(j).IndexImg = 0
                TabReglette(j).Lettre = ""
                Exit For
            End If
        Next
    Next
    
    For i = 1 To Len(StrLettresAChanger)
        If Mid(StrLettresAChanger, i, 1) = "?" Then
            STRPIOCHE = STRPIOCHE & "?"
        Else
            For j = 1 To Len(STRPIOCHE)
                If Asc(Mid(STRPIOCHE, j, 1)) > Asc(Mid(StrLettresAChanger, i, 1)) Or Mid(STRPIOCHE, j, 1) = "?" Then
                    STRPIOCHE = Left(STRPIOCHE, j - 1) & Mid(StrLettresAChanger, i, 1) & Right(STRPIOCHE, Len(STRPIOCHE) - j + 1)
                    Exit For
                End If
            Next
        End If
    Next
    
    ' Désactiver les boutons
    CmdPasser.Enabled = False
    CmdJouer.Enabled = False
    CmdRAZJetons.Enabled = False
    
    FillLettreJoueur
    
    ' Au tour du PC
    If MODEJEU = Normal Then
        CmdMelanger.Enabled = False
        TourPC
    End If
    
End Sub

Private Sub CmdFinTour_Click()

    DoEvents
    TimChrono.Enabled = False
    TimTourSuivant.Enabled = True
    
End Sub

Private Sub CmdJouer_Click()

    Dim ms As ClsMotScribble
    Dim i As Integer
    Dim j As Integer
    Dim Orientation As Integer
    Dim col As Integer
    Dim lig As Integer
    Dim x As Integer
    Dim y As Integer
    Dim Penalites As Integer
    Dim StrTemp  As String

    Set ms = CompterPoints
    
    If Asc(Left(ms.Pos, 1)) >= 65 Then
        Orientation = 0
    Else
        Orientation = 1
    End If
    
    col = G.GetColonne(ms.Pos)
    lig = G.GetLigne(ms.Pos)
    
    For i = 1 To Len(ms.Mot)
        ' Enlever la lettre au joueur
        For j = 0 To 6
            ' La lettre est sur la grille ?
            If PicLettreJoueur(j).Left + PicLettreJoueur(0).Width >= PicGrille.Left And PicLettreJoueur(j).Left < PicGrille.Left + PicGrille.Width Then
                If PicLettreJoueur(j).Top + PicLettreJoueur(0).Height >= PicGrille.Top And PicLettreJoueur(j).Top < PicGrille.Top + PicGrille.Height Then
                    ' Est-ce cette lettre la ?
                    x = (PicLettreJoueur(j).Left - PicGrille.Left) / ((PicGrille.Width - 2) / 15)
                    y = (PicLettreJoueur(j).Top - PicGrille.Top) / ((PicGrille.Height - 2) / 15)
            
                    If Orientation = 0 Then
                        If x = col + i - 1 And y = lig Then
                            ' Enlever la lettre
                            PicLettreJoueur(j).Tag = ""
                            Exit For
                        End If
                    Else
                        If x = col And y = lig + i - 1 Then
                            ' Enlever la lettre
                            PicLettreJoueur(j).Tag = ""
                            Exit For
                        End If
                    End If
                End If
            End If
        Next
    Next
    
    ' Désactiver les boutons
    CmdJouer.Enabled = False
    CmdRAZJetons.Enabled = False
    CmdMelanger.Enabled = False
    CmdPasser.Enabled = False
    CmdChanger.Enabled = False
    
    ' Désactiver les jetons pendant le tour de l'ordinateur (Normal/Duplicate)
    For i = 0 To 6
        PicLettreJoueur(i).Enabled = False
    Next
    
    ' Si premier tour, passer au deuxieme
    If G.IsPremierMot = True Then G.IsPremierMot = False
    
    ' Compter le nombre de Je passe
    NbJePasse = 0
    
    ' Duplicate
    If MODEJEU = Duplicate Then
        MsDuplicate.Mot = ms.Mot
        MsDuplicate.Pos = ms.Pos
        MsDuplicate.Pts = ms.Pts
        Exit Sub
    End If
    
    ' Ajouter au score du joueur
    TxtPtsJoueur = CStr(Val(TxtPtsJoueur) + ms.Pts)
    ' Afficher le coup dans l'historique
    AjouteHistorique "Joueur", ms
        
    StatusBar.Panels(2).Text = "0"
    
    ' Placer le mot sur la grille
    G.PlacerMot ms
    DrawGrille
    Set ms = Nothing
    
        
    ' Remettre des lettres
    FillLettreJoueur
    
    ' Vérifier si partie terminée...
    For i = 0 To 6
        If PicLettreJoueur(i).Tag <> "" Then
            i = 0
            Exit For
        End If
    Next
    
    ' la partie est termniée ?
    If i <> 0 Then
        ' Arrêter le chrono s'il est en route
        TimChrono.Enabled = False
        ' Compter les points restants du PC
        Penalites = 0
        For i = 1 To Len(LETTRESUC)
            Penalites = Penalites + G.PointLettres(LCase(Mid(LETTRESUC, i, 1)))
        Next
        
        StrTemp = "Joueur|" & TxtPtsJoueur & "|+ " & CStr(Penalites) & "|" & CStr(Val(TxtPtsJoueur) + Penalites)
        StrTemp = StrTemp & "|PC|" & TxtPts & "|- " & CStr(Penalites) & "|" & CStr(Val(TxtPts) - Penalites)
        
        TxtPts = CStr(Val(TxtPts) - Penalites)
        TxtPtsJoueur = CStr(Val(TxtPtsJoueur) + Penalites)
        
        Load FrmFin
        FrmFin.Affiche StrTemp
        FrmFin.Show vbModal
        
        MnuArreterLaPartie_Click
        
    Else
        If MODEJEU = Normal Then
            ' A l'ordinateur de jouer....
            Tour = PC
            Call TourPC
        Else
            ' Solo
            CmdPasser.Enabled = True
            For i = 0 To 6
                PicLettreJoueur(i).Enabled = True
            Next
            If Len(STRPIOCHE) >= 7 Then CmdChanger.Enabled = True
            FillLettreJoueur
        End If
    End If
    
End Sub

Private Sub TourPC()
    
    Dim MsMotRetenu As ClsMotScribble
    Dim NbMots As Long
    Dim StrMot As String
    Dim Pts As Integer
    Dim LgMot As Integer
    Dim T0 As Long
    Dim NbSol As Long
    Dim i As Integer
    Dim StrLettresJoueur As String
    Dim StrTemp As String
    Dim Penalites As Integer
    Dim BlnChange As Boolean
    
    Set MsMotRetenu = New ClsMotScribble
    Randomize (Time)
    
    
    ' Un sablier pour attendre....
    Screen.MousePointer = vbHourglass
    
    ' Et verouiller les lettres du joueur...
    For i = 0 To 6
        PicLettreJoueur(i).Enabled = False
    Next
    
    
    ' C'est le tour de l'ordinateur
    Tour = PC
    If PARTIECHRONOMETREE = "oui" Then InitChrono
    
    '  Chercher une solution....
    NbSol = G.TrouveSolution(LETTRESUC)
    TimChrono.Enabled = False
    
    Screen.MousePointer = vbNormal
    
    If NbSol = 0 Then
        ' Pas de solution !
        DrawGrille
        If Len(STRPIOCHE) >= 7 Then
            ' On peut encore changer des lettres
            ChangeNLettres
            NbJePasse = 0
        Else
            ' Il n'y a pas d'autre choix que de passer son tour
            Load FrmMsgbox
            FrmMsgbox.LblMsg = "L'ordinateur passe son tour !"
            FrmMsgbox.Show vbModal
            NbJePasse = NbJePasse + 1
            If NbJePasse >= 6 Then
                ' Si 3 fois passer / joueur, la partie est terminée...
                FinSurJePasse
                Exit Sub
            End If
        End If
    Else
        ' Prendre le mot correspondant auNIVEAU (de  1 à 5 )
        Set MsMotRetenu = G.GetSolution(Val(NIVEAU))
    
        ' L'ordinateur peux changer des lettres si le score est insuffisant
        BlnChange = False
        If Len(STRPIOCHE) >= 7 Then
            ' on peux changer ...
            If NIVEAU >= 3 Then
                ' à partir du niveau 3
                If MsMotRetenu.Pts < NIVEAU * 3 Then
                    ' car le score est insuffisant
                    BlnChange = True
                    ' Changer des lettres
                    ChangeNLettres
                    ' RAZ du nombre de 'Je passe'
                    NbJePasse = 0
                End If
            End If
        End If
        
        If BlnChange = False Then
        
            ' RAZ du nombre de 'Je passe'
            NbJePasse = 0
            
            ' On place le mot
            G.IsPremierMot = False
            DrawGrille
            AfficheMotEnBleu MsMotRetenu
            
            ' Enlever de LETTRESUC le mot placé
            EnleverLettre MsMotRetenu, G
            ' Le placer sur la grille
            G.PlacerMot MsMotRetenu
            ' Mettre à jour l'historique de partie et le score PC
            AjouteHistorique "PC", MsMotRetenu
            TxtPts = CStr(Val(TxtPts) + MsMotRetenu.Pts)
            ' une petite pause en cas ou...
            If PAUSEAPRESCOUP = "oui" Then
                PauseApresCoupPc
                If MnuArreterLaPartie.Visible = False Then Exit Sub
            End If
        End If
    End If
    
    ' Remettre des lettres à l'ordinateur
    FillLettreUC
    
    ' Fin de la partie ?
    If STRPIOCHE = "" And LETTRESUC = "" Then
        ' Compter les points restants du joueur (Penalites)
        For i = 0 To 6
            If PicLettreJoueur(i).Tag <> "" Then
                StrLettresJoueur = StrLettresJoueur & PicLettreJoueur(i).Tag
            End If
        Next
            
        Penalites = 0
        For i = 1 To Len(StrLettresJoueur)
            Penalites = Penalites + G.PointLettres(LCase(Mid(StrLettresJoueur, i, 1)))
        Next
        ' Prépare l'affichage du tableau
        StrTemp = "Joueur|" & TxtPtsJoueur & "|- " & CStr(Penalites) & "|" & CStr(Val(TxtPtsJoueur) - Penalites)
        StrTemp = StrTemp & "|PC|" & TxtPts & "|+ " & CStr(Penalites) & "|" & CStr(Val(TxtPts) + Penalites)
        ' Mettre à jour les scores
        TxtPts = CStr(Val(TxtPts) + Penalites)
        TxtPtsJoueur = CStr(Val(TxtPtsJoueur) - Penalites)
                
        Load FrmFin
        FrmFin.Affiche StrTemp
        FrmFin.Show vbModal
        
        ' Reafficher les menus..
        CmdPasser.Enabled = False
        MnuSolution.Enabled = False
        MnuArreterLaPartie.Visible = False
        MnuNouvellePartie.Visible = True
        MnuParamertes.Enabled = True
        MnuChargerPartie.Enabled = True
        ' Desactiver les jetons du joueur
        For i = 0 To 6
            PicLettreJoueur(i).Enabled = False
        Next
        
    Else
        ' Au tour du joueur ....
        Call TourJoueur
    End If
    
End Sub

Private Sub TourJoueur()

    Dim i As Integer
    
    ' Initialiser le chronomètre sio nécessaire
    If PARTIECHRONOMETREE = "oui" Then InitChrono
    Tour = Joueur
    
    ' Réactiver les lettres du joueur
    For i = 0 To 6
        PicLettreJoueur(i).Enabled = True
    Next
    
    ' Remettre le pointeur par défaut
    Screen.MousePointer = vbNormal
    
    ' Réactiver les boutons
    CmdPasser.Enabled = True
    CmdMelanger.Enabled = True
    If Len(STRPIOCHE) >= 7 Then
        CmdChanger.Enabled = True
    Else
        CmdChanger.Enabled = False
    End If

End Sub

Private Sub EnleverLettre(ms As ClsMotScribble, GBase As ClsGrille)

    Dim i As Integer
    Dim Car As String
    Dim p As Integer
    Dim StrMot As String
    
    StrMot = ms.Mot
    
    Dim col As Integer
    Dim lig As Integer
    Dim Orientation As Integer
    
    If Asc(Left(ms.Pos, 1)) >= 65 Then
        lig = Asc(Left(ms.Pos, 1)) - 65
        col = Val(Right(ms.Pos, Len(ms.Pos) - 1)) - 1
        Orientation = 1 ' H
    Else
        col = Val(Left(ms.Pos, Len(ms.Pos) - 1)) - 1
        lig = Asc(Right(ms.Pos, 1)) - 65
        Orientation = 2 ' V
    End If
    
    For i = 1 To Len(ms.Mot)
        
        If GBase.Cell(col, lig) = "" Then
            If Asc(Mid(ms.Mot, i, 1)) < 97 Then Mid(StrMot, i, 1) = "?"
            
            Car = UCase(Mid(StrMot, i, 1))
            
            p = InStr(LETTRESUC, Car)
            If p Then LETTRESUC = Left(LETTRESUC, p - 1) & Right(LETTRESUC, Len(LETTRESUC) - p)
        End If
        
        If Orientation = 1 Then
            col = col + 1
        Else
            lig = lig + 1
        End If
            
    Next

End Sub

Private Sub CmdMelanger_Click()

    Dim i As Integer
    Dim j As Integer
    Dim Top As Integer
    Dim TabMelange() As TMelange
    Dim TempMelange As TMelange
   
    ReDim TabMelange(0)
    For i = 0 To 12
        If TabReglette(i).Lettre <> "" Then
            Top = UBound(TabMelange) + 1
            ReDim Preserve TabMelange(Top)
            TabMelange(Top).IndexBloc = i
            TabMelange(Top).Lettre = TabReglette(i).Lettre
            TabMelange(Top).IndexImg = TabReglette(i).IndexImg
            TabMelange(Top).Left = PicLettreJoueur(TabMelange(Top).IndexImg).Left
        End If
    Next
        
    For i = 1 To UBound(TabMelange)
        j = Int(Rnd() * UBound(TabMelange)) + 1
        TempMelange = TabMelange(i)
        
        TabMelange(i).IndexImg = TabMelange(j).IndexImg
        TabMelange(i).Lettre = TabMelange(j).Lettre
        
        TabMelange(j).IndexImg = TempMelange.IndexImg
        TabMelange(j).Lettre = TempMelange.Lettre
        
    Next
        
    For i = 1 To UBound(TabMelange)
        TabReglette(TabMelange(i).IndexBloc).IndexImg = TabMelange(i).IndexImg
        TabReglette(TabMelange(i).IndexBloc).Lettre = TabMelange(i).Lettre
        PicLettreJoueur(TabMelange(i).IndexImg).Left = TabMelange(i).Left
    Next
        
        
    
End Sub

Private Sub CmdPasser_Click()

    Dim i As Integer
    Dim j As Integer
    Dim Pas As Integer
    Dim StrTemp As String
    Dim ms As ClsMotScribble
            
    ' Désactiver les bouton
    CmdJouer.Enabled = False
    CmdPasser.Enabled = False
    CmdRAZJetons.Enabled = False
    CmdMelanger.Enabled = False
    CmdChanger.Enabled = False
    
    NbJePasse = NbJePasse + 1
    If NbJePasse >= 6 Then
        FinSurJePasse
        Exit Sub
    End If
    
    If MODEJEU = Duplicate Then
        For i = 0 To 6
            PicLettreJoueur(i).Enabled = False
        Next
        MsDuplicate.Pts = -1
        Exit Sub
    Else
        ' Remettre toutes les lettres dans le TabReglette
        Pas = (PicReglette.Width - 3) / 13
        For i = 0 To 6
            If PicLettreJoueur(i).Tag <> "" Then
                ' Si la lettre se trouve sur la grille
                If PicLettreJoueur(i).Left + PicLettreJoueur(0).Width >= PicGrille.Left And PicLettreJoueur(i).Left < PicGrille.Left + PicGrille.Width Then
                    If PicLettreJoueur(i).Top + PicLettreJoueur(0).Height >= PicGrille.Top And PicLettreJoueur(i).Top < PicGrille.Top + PicGrille.Height Then
                        For j = 3 To 12
                            If TabReglette(j).Lettre = "" Then
                                PicLettreJoueur(i).Left = PicReglette.Left + j * Pas + 2
                                PicLettreJoueur(i).Top = PicReglette.Top + 2
                                TabReglette(j).Lettre = PicLettreJoueur(i).Tag
                                TabReglette(j).IndexImg = i
                                Exit For
                            End If
                        Next
                    End If
                End If
            End If
        Next
        
        DoEvents
        
        If MODEJEU = Normal Then
            ' A l'ordinateur de jouer....
            Call TourPC
        Else
            ' En solo fin de partie
            StrTemp = "Joueur|" & TxtPtsJoueur & "|0|" & CStr(Val(TxtPtsJoueur))
            StrTemp = StrTemp & "|-|-|-|-"
            Load FrmFin
            FrmFin.Affiche StrTemp
            FrmFin.Show vbModal
    
            MnuArreterLaPartie_Click

        End If
    End If

End Sub

Private Sub CmdRAZJetons_Click()

    Dim i As Integer
    Dim j As Integer
    Dim Pas As Integer
    
    Pas = (PicReglette.Width - 3) / 13
    
    For i = 0 To 6
        If PicLettreJoueur(i).Tag <> "" Then
            ' L'image de la lettre est sur la grille ?
            If PicLettreJoueur(i).Left + PicLettreJoueur(0).Width >= PicGrille.Left And PicLettreJoueur(i).Left < PicGrille.Left + PicGrille.Width Then
                If PicLettreJoueur(i).Top + PicLettreJoueur(0).Height >= PicGrille.Top And PicLettreJoueur(i).Top < PicGrille.Top + PicGrille.Height Then
                    For j = 3 To 12
                        If TabReglette(j).Lettre = "" Then
                            PicLettreJoueur(i).Left = PicReglette.Left + j * Pas + 2
                            PicLettreJoueur(i).Top = PicReglette.Top + 2
                            TabReglette(j).Lettre = PicLettreJoueur(i).Tag
                            TabReglette(j).IndexImg = i
                            ' Si jocker, réafficher un blanc
                            If Asc(PicLettreJoueur(i).Tag) < 97 Then
                                PicLettreJoueur(i).PaintPicture FrmPics.PicLettre(26).Image, 0, 0
                                PicLettreJoueur(i).Tag = "?"
                            End If
                            Exit For
                        End If
                    Next
                End If
            End If
        End If
    Next
    
    ' Désactiver le bouton mélanger si la reglette est vide .
    CmdMelanger.Enabled = True
    i = 0
    For j = 0 To 12
        If TabReglette(j).Lettre <> "" Then i = i + 1
    Next
    If i < 2 Then CmdMelanger.Enabled = False
    
    ' Désactiver les boutons
    CmdJouer.Enabled = False
    CmdRAZJetons.Enabled = False
    StatusBar.Panels(2).Text = ""

End Sub


Private Sub Form_Load()
    
    Dim StrDate As String
    Dim d As Integer
    Dim SecretKey As String
    
    ' Afficher le splash screen
    FrmSplash.Show
    DoEvents
    
    ' Initiliser les images des jetons, de la grille et de la reglette en fonction de la résolution
    InitAffichage
    
    ' Lire les valeurs dans le registre
    NIVEAU = RegGetDword(HKEY_LOCAL_MACHINE, KEYSCRIBBLE, "Niveau")
    CHRONOMETRE = RegGetDword(HKEY_LOCAL_MACHINE, KEYSCRIBBLE, "Chrono")
    PARTIECHRONOMETREE = RegGetString(HKEY_LOCAL_MACHINE, KEYSCRIBBLE, "PartieChronometree")
    PREMIERTOUR = RegGetDword(HKEY_LOCAL_MACHINE, KEYSCRIBBLE, "PremierTour") '0 = Aleatoire par défaut
    MODEJEU = RegGetDword(HKEY_LOCAL_MACHINE, KEYSCRIBBLE, "ModeJeu") '0
    StrDico = RegGetString(HKEY_LOCAL_MACHINE, KEYSCRIBBLE, "Dico")
    PAUSEAPRESCOUP = RegGetString(HKEY_LOCAL_MACHINE, KEYSCRIBBLE, "PauseApresCoup")
    TEMPSPAUSE = RegGetDword(HKEY_LOCAL_MACHINE, KEYSCRIBBLE, "TempsPause")
    
    ' ou prendre la valeur par défaut....
    If PREMIERTOUR > 2 Then PREMIERTOUR = Aleatoire
    If NIVEAU = 0 Or NIVEAU > 5 Then NIVEAU = 3
    If MODEJEU < 0 Or MODEJEU > 2 Then MODEJEU = 0
    If StrDico = "" Then StrDico = "ODS"
    If CHRONOMETRE = 0 Then CHRONOMETRE = 120
    If PARTIECHRONOMETREE = "" Then PARTIECHRONOMETREE = "non"      ' non........
    If PAUSEAPRESCOUP = "" Then PAUSEAPRESCOUP = "oui"              ' Par défaut une pause apres le coup du PC
    If TEMPSPAUSE = 0 Then TEMPSPAUSE = 5                           ' 5 secondes de pause apres le coup du PC
    
            
    Set Dico = New ClsDictionnaire
    If StrDico = "" Then StrDico = "ODS"
    If Dico.InitDictionnaire(StrDico) = False Then
        MsgBox "Impossible d'initialiser le dictionnaire"
    End If
    Me.Caption = "  Megafan Scribble - " & StrDico
        
    Set Ana = New ClsMots
    Ana.Dictionnaire = Dico
    
    
    ' Afficher une grille vide dans le fond
    Set G = New ClsGrille
    DrawGrille
    Set G = Nothing
        
    PicOK.Visible = False
    PicKO.Visible = False
    
    ' Initialiser la grille de l'historique
    GridHistorique.FormatString = "^N°|^Nom|^Mot|^Pos|^Pts"
    GridHistorique.ColWidth(0) = 350
    GridHistorique.ColWidth(1) = 850
    GridHistorique.ColWidth(2) = 1400
    GridHistorique.ColWidth(3) = 600
    GridHistorique.ColWidth(4) = 600
    GridHistorique.RowHeight(1) = 0
    
    Unload FrmSplash
    
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)

    End
    
End Sub


Private Sub Mnu1erTour_Click()

    Frm1erTour.Show vbModal

End Sub

Private Sub MnuAnagrammes_Click()

    FrmAnagrammes.Show vbModal
    
End Sub

Private Sub MnuAPropos_Click()

    FrmAProposDe.Show vbModal

End Sub

Private Sub MnuArreterLaPartie_Click()

    Dim i As Integer
    'MsgBox "Etes-vous certain de vouloir arrêter cette partie ?"
    
    TimChrono.Enabled = False
    MnuArreterLaPartie.Visible = False
    MnuSolution.Enabled = False
    DoEvents
    If MODEJEU = Duplicate Then
        If GDuplicate.IsSearching Then GDuplicate.Abort = True
    Else
        If G.IsSearching Then G.Abort = True
    End If
    
    For i = 0 To 6
        PicLettreJoueur(i).Enabled = False
    Next
    
    MnuNouvellePartie.Visible = True
    MnuParamertes.Enabled = True
    MnuChargerPartie.Enabled = True
    
    CmdJouer.Enabled = False
    CmdPasser.Enabled = False
    CmdMelanger.Enabled = False
    CmdRAZJetons.Enabled = False
    CmdChanger.Enabled = False
    
    LblChrono.ForeColor = vbBlack
    LblChrono.BackColor = &H80FFFF
    LblChrono.Caption = "00:00"

End Sub

Private Sub MnuChangerDico_Click()
        
    Dim StrNewDico As String
    
    FrmDictionnaire.Show vbModal
    StrNewDico = RegGetString(HKEY_LOCAL_MACHINE, KEYSCRIBBLE, "Dico")
    
    If StrNewDico <> StrDico Then
        FrmSplash.Show
        DoEvents
        Dico.InitDictionnaire (RegGetString(HKEY_LOCAL_MACHINE, KEYSCRIBBLE, "Dico"))
        Unload FrmSplash
        FrmMain.Caption = "  Megafan Scribble - " & StrNewDico
    End If
   
End Sub

Private Sub MnuModeDeJeu_Click()

    FrmModeDeJeu.Show vbModal
    
End Sub

Private Sub MnuNiveauPC_Click()

    FrmNiveaux.Show vbModal

End Sub

Private Sub MnuNouvellePartie_Click()
    
    Dim StrP As String
    Dim i As Integer
    Dim j As Integer
    
    MnuNouvellePartie.Visible = False
    MnuArreterLaPartie.Visible = True
    MnuParamertes.Enabled = False
    MnuChargerPartie.Enabled = False
    NbJePasse = 0
    
    ' Un peu de hasard
    Randomize Time
        
    ' Effacer les lettres du joueur
    For i = 0 To 6
        PicLettreJoueur(i).Tag = ""
    Next
    ' Vider la reglette
    For i = 0 To 12
        TabReglette(i).Lettre = ""
        TabReglette(i).IndexImg = -1
    Next
    ' Effacer les lettres de l'ordinateur
    LETTRESUC = ""
    
    ' Initialiser la pioche (en fonction du dictionnaire)
    StrP = STRPIOCHEDICO
    
    'StrP = "04A01B01C01D04E01F02I02L02M02N02O02P02S02T03U01V02?"
    
    STRPIOCHE = ""
    For i = 1 To Len(StrP) Step 3
        j = Val(Mid(StrP, i, 2))
        STRPIOCHE = STRPIOCHE & String(j, Mid(StrP, i + 2, 1))
    Next
    
    ' Effacer la grille d'historique de partie
    GridHistorique.Rows = 2
    GridHistorique.RowHeight(1) = 0
    
    ' Effacer les scores
    TxtPts = "0"
    TxtPtsJoueur = "0"
    
    Set G = New ClsGrille
    DrawGrille
    
    If MODEJEU = Normal Then
        Tour = -1
        If PREMIERTOUR = Aleatoire Then
            Tour = Int(Rnd() * 2) + 1
        End If
        
        If PREMIERTOUR = Joueur Or Tour = Joueur Then
            Tour = Joueur
            FillLettreJoueur
            MnuSolution.Enabled = True
            FillLettreUC
            If PARTIECHRONOMETREE = "oui" Then InitChrono
            TourJoueur
        Else
            Tour = PC
            FillLettreUC
            FillLettreJoueur
            CmdChanger.Enabled = False
            MnuSolution.Enabled = True
            If PARTIECHRONOMETREE = "oui" Then InitChrono
            TourPC
        End If
    End If
    
    If MODEJEU = Duplicate Then
        
        Set GDuplicate = New ClsGrille
        
        ' Duplicate
        LETTRESUC = ""
        DuplicateAjouteLettres
        CmdPasser.Enabled = True
        CmdMelanger.Enabled = True
        
        Set MsDuplicate = New ClsMotScribble
        InitChrono
        GDuplicate.TrouveSolution LETTRESUC
        
    End If
    
    If MODEJEU = Solo Then
        
        FillLettreJoueur
        MnuSolution.Enabled = True
        TourJoueur
        
    End If

End Sub

Private Sub MnuQuitter_Click()
    
    End

End Sub

Private Sub MnuSolution_Click()

    Dim i As Integer
    Dim LettresJoueur As String
    Dim MsS As New ClsMotScribble
    Dim NbSol As Long
    
    If G Is Nothing Then Exit Sub
    
    Me.MousePointer = vbHourglass
    For i = 0 To 6
        If PicLettreJoueur(i).Tag <> "" Then
            If Asc(PicLettreJoueur(i).Tag) < 97 Then
                LettresJoueur = LettresJoueur & "?"
            Else
                LettresJoueur = LettresJoueur & PicLettreJoueur(i).Tag
            End If
        End If
    Next
    
    NbSol = G.TrouveSolution(LettresJoueur)
            
    Load FrmMsgbox
    If NbSol = 0 Then
        FrmMsgbox.LblMsg = "Il n'y a pas de solution !"
    Else
        Set MsS = G.GetSolution(5)
        FrmMsgbox.LblMsg = "Meilleure solution : " & MsS.Mot & " (" & MsS.Pos & ") pour " & MsS.Pts & " pts."
    End If
    Me.MousePointer = vbNormal
    FrmMsgbox.Show vbModal
    
    Set MsS = Nothing

End Sub

Private Sub PicLettreJoueur_MouseDown(Index As Integer, Button As Integer, Shift As Integer, x As Single, y As Single)

    Dim p As POINTAPI
    GetCursorPos p

    If Button = 1 Then
        ' Clic Gauche, déplacement d'un Jeton
        LettreJoueur.Idx = Index
        LettreJoueur.Dx = x - (Me.Left / Screen.TwipsPerPixelX + PicLettreJoueur(Index).Left + x - p.x)
        LettreJoueur.Dy = y - (Me.Top / Screen.TwipsPerPixelY + PicLettreJoueur(Index).Top + y - p.y)
        LettreJoueur.OldX = PicLettreJoueur(Index).Left
        LettreJoueur.OldY = PicLettreJoueur(Index).Top
        LettreJoueur.BorderDiffX = PicLettreJoueur(Index).Width - (Me.Left / Screen.TwipsPerPixelX + PicLettreJoueur(Index).Left + x - p.x)
        LettreJoueur.BorderDiffY = PicLettreJoueur(Index).Height - (Me.Top / Screen.TwipsPerPixelY + PicLettreJoueur(Index).Top + y - p.y)
        PicLettreJoueur(Index).ZOrder
        
        ' Départ depuis la reglette ?
        If PicLettreJoueur(Index).Top > PicGrille.Top + PicGrille.Height Then
            LettreJoueur.PosBlocLettre = CInt(PicLettreJoueur(Index).Left - PicReglette.Left) / ((PicReglette.Width - 2) / 13)
        Else
            LettreJoueur.PosBlocLettre = -1
        End If
        
        TimDrag.Enabled = True
    Else
    
    
    End If

End Sub

Private Sub PicLettreJoueur_MouseUp(Index As Integer, Button As Integer, Shift As Integer, x As Single, y As Single)
    
    Dim Ligne As Double
    Dim Colonne As Double
    Dim i As Integer
    Dim j As Integer
    Dim Pas As Integer
    Dim ms As ClsMotScribble
    
    If Button <> 1 Then Exit Sub
    TimDrag.Enabled = False
    
    Pas = (PicReglette.Width - 3) / 13
    
    ' 1°) sur la grille ?
    If PicLettreJoueur(LettreJoueur.Idx).Left + PicLettreJoueur(0).Width >= PicGrille.Left And PicLettreJoueur(LettreJoueur.Idx).Left < PicGrille.Left + PicGrille.Width Then
        If PicLettreJoueur(LettreJoueur.Idx).Top + PicLettreJoueur(0).Height >= PicGrille.Top And PicLettreJoueur(LettreJoueur.Idx).Top < PicGrille.Top + PicGrille.Height Then
            Ligne = (PicLettreJoueur(LettreJoueur.Idx).Top - PicGrille.Top) / ((PicGrille.Height - 2) / 15)
            
            ' Déterminer la ligne
            If Ligne - Int(Ligne) < 0.6 Then
                Ligne = Int(Ligne)
            Else
                Ligne = Int(Ligne) + 1
            End If
            If Ligne > 14 Then Ligne = 14
            If Ligne < 0 Then Ligne = 0
            
            ' Déterminer la colonne
            Colonne = (PicLettreJoueur(LettreJoueur.Idx).Left - PicGrille.Left) / ((PicGrille.Width - 2) / 15)
            If Colonne - Int(Colonne) < 0.6 Then
                Colonne = Int(Colonne)
            Else
                Colonne = Int(Colonne) + 1
            End If
            If Colonne > 14 Then Colonne = 14
            If Colonne < 0 Then Colonne = 0
        
            If G.Cell(CInt(Colonne), CInt(Ligne)) = "" Then
                PicLettreJoueur(Index).Left = CInt(PicGrille.Left + Colonne * (PicGrille.Width - 2) / 15 + 2)
                PicLettreJoueur(Index).Top = CInt(PicGrille.Top + Ligne * (PicGrille.Height - 2) / 15 + 2)
                
                ' Jeton déposé sur un autre (but : inverser les 2)
                For i = 0 To 6
                    If i <> Index Then
                        If PicLettreJoueur(i).Left = PicLettreJoueur(Index).Left And PicLettreJoueur(i).Top = PicLettreJoueur(Index).Top Then
                            ' Inverser les 2 jetons
                            PicLettreJoueur(i).Left = LettreJoueur.OldX
                            PicLettreJoueur(i).Top = LettreJoueur.OldY
                            
                            If LettreJoueur.PosBlocLettre >= 0 Then
                                ' le Jeton proviens de la reglette
                                If Asc(PicLettreJoueur(i).Tag) < 97 Then
                                    ' C'est un jocker qui est 'ecrasé'
                                    PicLettreJoueur(i).Tag = "?"
                                    PicLettreJoueur(i).PaintPicture FrmPics.PicLettre(26).Image, 0, 0
                                End If
                                ' inverser les references des 2 jetons dans la reglette
                                TabReglette(LettreJoueur.PosBlocLettre).Lettre = PicLettreJoueur(i).Tag
                                TabReglette(LettreJoueur.PosBlocLettre).IndexImg = i
                            End If
                        End If
                    End If
                Next
                
                ' Demander lettre si Jocker
                If PicLettreJoueur(Index).Tag = "?" Then
                    FrmJocker.Show vbModal
                    If FrmJocker.TxtLettre = "" Then Exit Sub
                    PicLettreJoueur(Index).Tag = FrmJocker.TxtLettre
                    Unload FrmJocker
                    CreerPicJocker (PicLettreJoueur(Index).Tag)
                    PicLettreJoueur(Index).PaintPicture FrmPics.PicTemp.Image, 0, 0
                End If
                
                ' Activer le bouton RAZ Jetons
                CmdRAZJetons.Enabled = True
                
                ' libérer la place
                If i = 7 Then
                    If LettreJoueur.PosBlocLettre >= 0 Then
                        TabReglette(LettreJoueur.PosBlocLettre).Lettre = ""
                        TabReglette(LettreJoueur.PosBlocLettre).IndexImg = -1
                        
                        ' Désactiver le bouton mélanger si la reglette est vide .
                        i = 0
                        For j = 0 To 12
                            If TabReglette(j).Lettre <> "" Then i = i + 1
                        Next
                        If i < 2 Then CmdMelanger.Enabled = False
                    End If
                    
                    ' Compter les points du coup.
                    Set ms = CompterPoints
                    StatusBar.Panels(2).Text = ms.Pts
                    Exit Sub
                End If
                
            End If
        End If
    End If
    
    ' 2°) sur la reglette ?
    If PicLettreJoueur(LettreJoueur.Idx).Left + PicLettreJoueur(0).Width >= PicReglette.Left And PicLettreJoueur(LettreJoueur.Idx).Left < PicReglette.Left + PicReglette.Width Then
        If PicLettreJoueur(LettreJoueur.Idx).Top + PicLettreJoueur(0).Height >= PicReglette.Top And PicLettreJoueur(LettreJoueur.Idx).Top < PicReglette.Top + PicReglette.Height Then
            Colonne = (PicLettreJoueur(LettreJoueur.Idx).Left - PicReglette.Left) / Pas
            If Colonne - Int(Colonne) < 0.6 Then
                Colonne = Int(Colonne)
            Else
                Colonne = Int(Colonne) + 1
            End If
            If Colonne > 12 Then Colonne = 12
            If Colonne < 0 Then Colonne = 0
            
            ' Sur une case vide
            If TabReglette(Colonne).Lettre = "" Then
                PicLettreJoueur(Index).Left = PicReglette.Left + Colonne * Pas + 2
                PicLettreJoueur(Index).Top = PicReglette.Top + 2
                TabReglette(Colonne).Lettre = PicLettreJoueur(Index).Tag
                TabReglette(Colonne).IndexImg = Index
                ' libérer la place
                If LettreJoueur.PosBlocLettre >= 0 Then
                    TabReglette(LettreJoueur.PosBlocLettre).Lettre = ""
                    TabReglette(LettreJoueur.PosBlocLettre).IndexImg = -1
                End If
                
                ' Désactiver le bouton mélanger si la reglette est vide .
                CmdMelanger.Enabled = True
                i = 0
                For j = 0 To 12
                    If TabReglette(j).Lettre <> "" Then i = i + 1
                Next
                If i < 2 Then CmdMelanger.Enabled = False
            Else
                PicLettreJoueur(Index).Left = PicReglette.Left + Colonne * Pas + 2
                PicLettreJoueur(Index).Top = PicReglette.Top + 2
                TabReglette(Colonne).Lettre = PicLettreJoueur(Index).Tag
                TabReglette(Colonne).IndexImg = Index
                ' Quelle est la lettre ecrasée
                For i = 0 To 6
                    If i <> Index Then
                        If PicLettreJoueur(i).Left = PicLettreJoueur(Index).Left And PicLettreJoueur(i).Top = PicLettreJoueur(Index).Top Then
                            ' Trouver une place pour l'ancienne lettre
                            If LettreJoueur.PosBlocLettre >= 0 Then
                                TabReglette(LettreJoueur.PosBlocLettre).Lettre = PicLettreJoueur(i).Tag
                                TabReglette(LettreJoueur.PosBlocLettre).IndexImg = i
                                PicLettreJoueur(i).Left = PicReglette.Left + LettreJoueur.PosBlocLettre * Pas + 2
                                PicLettreJoueur(i).Top = PicReglette.Top + 2
                                Exit For
                            Else
                                ' Une place libre
                                For j = 3 To 12
                                    If TabReglette(j).Lettre = "" Then
                                        TabReglette(j).Lettre = PicLettreJoueur(i).Tag
                                        TabReglette(j).IndexImg = i
                                        PicLettreJoueur(i).Left = PicReglette.Left + j * Pas + 2
                                        PicLettreJoueur(i).Top = PicReglette.Top + 2
                                        Exit For
                                    End If
                                Next
                                Exit For
                            End If
                        End If
                    End If
                Next
            End If
                
            ' Si toutes les lettres sont sur PicGrille, desactiver le bouton [RAZ Jerons]
            For i = 0 To 6
                If PicLettreJoueur(i).Tag <> "" Then
                    If PicLettreJoueur(i).Left + PicLettreJoueur(0).Width >= PicGrille.Left And PicLettreJoueur(i).Left < PicGrille.Left + PicGrille.Width Then
                        If PicLettreJoueur(i).Top + PicLettreJoueur(0).Height >= PicGrille.Top And PicLettreJoueur(i).Top < PicGrille.Top + PicGrille.Height Then
                            Exit For
                        End If
                    End If
                End If
            Next
            
            If i = 7 Then CmdRAZJetons.Enabled = False
                
            ' Compter les points
            Set ms = CompterPoints
            StatusBar.Panels(2).Text = ms.Pts
            
            ' Si jocker, réafficher un blanc
            If Asc(PicLettreJoueur(Index).Tag) < 97 Then
                PicLettreJoueur(Index).PaintPicture FrmPics.PicLettre(26).Image, 0, 0
                PicLettreJoueur(Index).Tag = "?"
            End If
            
            Exit Sub
        End If
    End If
    
    ' Finalement la remettre à l'ancienne position.
    PicLettreJoueur(LettreJoueur.Idx).Left = LettreJoueur.OldX
    PicLettreJoueur(LettreJoueur.Idx).Top = LettreJoueur.OldY

End Sub

Private Sub TimChrono_Timer()

    Dim LngTemps As Long
    
    LngTemps = CHRONOMETRE - CLng((GetTickCount() - CLng(TimChrono.Tag)) / 1000)

    ' Int et pas Cint !
    If PARTIECHRONOMETREE = "oui" Then
        LblChrono = Format(Int(LngTemps / 60), "00") & ":" & Format(LngTemps Mod 60, "00")
    
        If LngTemps <= 30 Then
            LblChrono.ForeColor = vbRed
        End If
    End If
   
    If MODEJEU = Duplicate Then
        If GDuplicate.IsSolutionTrouvee = True Then
            ' Attendre que le joueur joue (Passer = -1)
            If MsDuplicate.Pts <> 0 Or MsDuplicate.Mot <> "" Then
                Screen.MousePointer = vbNormal
                If PARTIECHRONOMETREE = "non" Then
                    CmdFinTour_Click
                    Exit Sub
                End If
                CmdFinTour.Enabled = True
            End If
        Else
            If MsDuplicate.Pts <> 0 Or MsDuplicate.Mot <> "" Then
                Screen.MousePointer = vbHourglass
            End If
        End If
             
        ' 3 secondes avant la fin, abandonner la recherche de l'odinateur
        If LngTemps <= 3 Then
            If GDuplicate.IsSolutionTrouvee = False Then
                GDuplicate.Abort = True
            End If
        End If
    Else
        ' Mode normal...
        If Tour = PC Then
            ' 3 secondes avant la fin, abandonner la recherche de l'odinateur
            If LngTemps <= 3 Then
                If G.IsSolutionTrouvee = False Then
                    G.Abort = True
                End If
            End If
        End If
    End If
            
    If LngTemps <= 0 Then
        If PARTIECHRONOMETREE = "oui" Then
            LblChrono = "00:00"
            DoEvents
            TimChrono.Enabled = False
            
            ' Declencher la suite en 'asynchrone'
            TimTourSuivant.Enabled = True
        End If
    End If
    
End Sub

Private Sub TimDrag_Timer()

    Dim p As POINTAPI
    Dim NewX As Integer
    Dim NewY As Integer
    
    GetCursorPos p
    NewX = p.x - (Me.Left / Screen.TwipsPerPixelX) - LettreJoueur.Dx
    NewY = p.y - (Me.Top / Screen.TwipsPerPixelY) - LettreJoueur.Dy
    
    
    ' Vérifier sortie de fenêtre
    If NewX < 0 Then NewX = 0
    If NewY < 0 Then NewY = 0
    If NewX >= Me.Width / Screen.TwipsPerPixelX - LettreJoueur.BorderDiffX Then NewX = Me.Width / Screen.TwipsPerPixelX - LettreJoueur.BorderDiffX
    If NewY >= Me.Height / Screen.TwipsPerPixelY - LettreJoueur.BorderDiffY Then NewY = Me.Height / Screen.TwipsPerPixelY - LettreJoueur.BorderDiffY
    
    
    PicLettreJoueur(LettreJoueur.Idx).Left = NewX
    PicLettreJoueur(LettreJoueur.Idx).Top = NewY

End Sub

Private Sub TimTourSuivant_Timer()

    ' Declencher en asynchrone cette fonction
    TimTourSuivant.Enabled = False
    If MODEJEU = Duplicate Then
        CmdFinTour.Enabled = False
        DuplicateTourSuivant
    Else
        CmdPasser_Click
    End If
    
End Sub

Private Sub TxtMotAVerifier_Change()

    PicKO.Visible = False
    PicOK.Visible = False
    If TxtMotAVerifier = "" Then Exit Sub
    
    If Ana.IsMotexiste(TxtMotAVerifier) Then
        PicOK.Visible = True
    Else
        PicKO.Visible = True
    End If
    
End Sub

Public Sub DrawGrille()

    Dim i As Integer
    Dim j As Integer
    Dim Car As String
    
    PicGrille.PaintPicture FrmPics.Grille.Image, 0, 0, PicGrille.Width, PicGrille.Height
       
    ' Afficher les lettres
    For j = 0 To 14
        For i = 0 To 14
            Car = G.Cell(i, j)
            If Car <> "" Then
                If Asc(Car) < 97 Then
                    ' C'est un Jocker, il faut afficher la lettre en rouge
                    CreerPicJocker (Car)
                    PicGrille.PaintPicture FrmPics.PicTemp.Image, i * (FrmPics.PicLettre(0).Width + 1) + 2, j * (FrmPics.PicLettre(0).Height + 1) + 2
                Else
                    PicGrille.PaintPicture FrmPics.PicLettre(Asc(Car) - 97).Image, i * (FrmPics.PicLettre(0).Width + 1) + 2, j * (FrmPics.PicLettre(0).Height + 1) + 2, TailleJetons, TailleJetons, 0, 0, TailleJetons, TailleJetons
                End If
            End If
        Next
    Next

    PicGrille.Refresh
        
End Sub

Private Sub FillLettreUC()

    Dim i As Integer
    Dim Pos As Integer
    Dim Car As String

    Randomize (Time)
    If Len(LETTRESUC) < 7 Then
        i = 1
        Do
            Pos = Int(Rnd() * Len(STRPIOCHE)) + 1
            Car = Mid(STRPIOCHE, Pos, 1)
            If STRPIOCHE <> "" Then STRPIOCHE = Left(STRPIOCHE, Pos - 1) & Right(STRPIOCHE, Len(STRPIOCHE) - Pos)
            LETTRESUC = LETTRESUC + Car
            If STRPIOCHE = "" Then Exit Do
            If Len(LETTRESUC) = 7 Then Exit Do
            i = i + 1
            If i = 8 Then Exit Do
        Loop While 1
    End If

    If Me.TextWidth(STRPIOCHE) >= StatusBar.Panels(1).Width Then
        StatusBar.Panels(1).Text = CStr(Len(STRPIOCHE)) & " lettres."
    Else
        StatusBar.Panels(1).Text = STRPIOCHE
    End If

End Sub

Private Sub CreerPicJocker(Car As String)

    Dim i As Integer
    Dim j As Integer
    Dim col As Long
    Dim r As Integer
    Dim v As Integer
    Dim b As Integer
    Dim cl As Long
    
    ' Copie de la lettre entière
    FrmPics.PicTemp.PaintPicture FrmPics.PicLettre(Asc(Car) - 65).Image, 0, 0
    ' Copie de la partie 'point' vide
    FrmPics.PicTemp.PaintPicture FrmPics.PicLettre(26).Image, 30, 32, , , 30, 32
    FrmPics.PicTemp.Refresh
    
    For j = 0 To FrmPics.PicTemp.Height
        For i = 0 To FrmPics.PicTemp.Width
            col = GetPixel(FrmPics.PicTemp.hDC, i, j)
            r = col And &HFF&
            v = (col \ 256) And &HFF&
            b = (col \ 65536) And &HFF&
            If r + v + b < 528 Then
                FrmPics.PicTemp.PSet (i, j), RGB(200, v, b)
            End If
        Next
    Next

End Sub

Private Sub AfficheMotEnBleu(MsMot As ClsMotScribble)

    Dim Car As String
    Dim i As Integer
    Dim j As Integer
    Dim k As Integer
    Dim StrPosition As String
    Dim Orientation As Integer
    
    StrPosition = MsMot.Pos
    If Asc(Left(StrPosition, 1)) >= Asc("A") And Asc(Left(StrPosition, 1)) <= Asc("O") Then
        Orientation = 1
        j = Asc(Left(StrPosition, 1)) - 65
        StrPosition = Right(StrPosition, Len(StrPosition) - 1)
        i = Val(StrPosition) - 1
    Else
        Orientation = 2
        j = Asc(Right(StrPosition, 1)) - 65
        StrPosition = Left(StrPosition, Len(StrPosition) - 1)
        i = Val(StrPosition) - 1
    End If
    
    For k = 1 To Len(MsMot.Mot)
        Car = Mid(MsMot.Mot, k, 1)
        
        If G.Cell(i, j) = "" Then
            If Car <> "" Then
                If Asc(Car) < 97 Then
                    CreerPicJocker (UCase(Car))
                Else
                    CreerLettreBleue (UCase(Car))
                End If
            
                PicGrille.PaintPicture FrmPics.PicTemp.Image, i * (FrmPics.PicLettre(0).Width + 1) + 2, j * (FrmPics.PicLettre(0).Height + 1) + 2
                'PicGrille.Refresh
                
            End If
        End If
        If Orientation = 1 Then
            i = i + 1
        Else
            j = j + 1
        End If
    Next
    DoEvents
    
End Sub

Private Sub CreerLettreBleue(Car As String)

    Dim i As Integer
    Dim j As Integer
    Dim col As Long
    Dim r As Integer
    Dim v As Integer
    Dim b As Integer
    Dim cl As Long
    
    ' Copie de la lettre entiére
    FrmPics.PicTemp.PaintPicture FrmPics.PicLettre(Asc(Car) - 65).Image, 0, 0
    
    For j = 0 To FrmPics.PicTemp.Height
        For i = 0 To FrmPics.PicTemp.Width
            col = GetPixel(FrmPics.PicTemp.hDC, i, j)
            r = col And &HFF&
            v = (col \ 256) And &HFF&
            b = (col \ 65536) And &HFF&
            If r + v + b < 528 Then
                FrmPics.PicTemp.PSet (i, j), RGB(r, v, 200)
            End If
        Next
    Next

End Sub

Private Sub FillLettreJoueur()

    Dim i As Integer
    Dim j As Integer
    Dim Pos As Integer
    Dim Car As String
    Dim Pas As Integer

    Randomize (Time)
    Pas = (PicReglette.Width - 3) / 13
    
    If Pas <> Int(Pas) Then
        MsgBox "Erreur dans le calcul du pas de l'image PicReglette", vbCritical
        End
    End If
        
    For i = 0 To 6
        If PicLettreJoueur(i).Tag = "" Then
            Pos = Int(Rnd() * Len(STRPIOCHE)) + 1
            Car = Mid(STRPIOCHE, Pos, 1)
            If Car = "" Then Exit For
            If STRPIOCHE <> "" Then STRPIOCHE = Left(STRPIOCHE, Pos - 1) & Right(STRPIOCHE, Len(STRPIOCHE) - Pos)
                        
            If Car <> "?" Then
                PicLettreJoueur(i).PaintPicture FrmPics.PicLettre(Asc(Car) - 65).Image, 0, 0
            Else
                PicLettreJoueur(i).PaintPicture FrmPics.PicLettre(26).Image, 0, 0
            End If
            
            For j = 3 To 12
                If TabReglette(j).Lettre = "" Then
                    PicLettreJoueur(i).Tag = LCase(Car)
                    PicLettreJoueur(i).Visible = True
                    PicLettreJoueur(i).Left = PicReglette.Left + j * Pas + 2
                    PicLettreJoueur(i).Top = PicReglette.Top + 2
                    TabReglette(j).Lettre = LCase(Car)
                    TabReglette(j).IndexImg = i
                    Exit For
                End If
            Next
            
            If STRPIOCHE = "" Then Exit For
        End If
    Next
    
    ' Masquer les pions non utilisés
    For i = 0 To 6
        If PicLettreJoueur(i).Tag = "" Then
            PicLettreJoueur(i).Visible = False
        Else
            PicLettreJoueur(i).Visible = True
        End If
    Next
    
    Me.Font.Size = StatusBar.Font.Size
    If Me.TextWidth(STRPIOCHE) >= StatusBar.Panels(1).Width Then
        StatusBar.Panels(1).Text = CStr(Len(STRPIOCHE)) & " lettres."
    Else
        StatusBar.Panels(1).Text = STRPIOCHE
    End If

End Sub

Private Sub InitAffichage()

    Dim i As Integer
        
    FrmMain.ForeColor = &H80FFFF
    
    Select Case Screen.Height / Screen.TwipsPerPixelY
        Case Is <= 600
            MsgBox "Résolution non géree." & vbCrLf & "Le jeu n'accepte qu'une résolution minimale de 1024 x 768 points.", vbCritical
            End
        Case Is <= 768
            ' 1024 x 768
            TailleJetons = 35
            FrmMain.Font.Size = 11
            FrmMain.StatusBar.Font.Size = 9
            Me.Height = Screen.Height - (38 * Screen.TwipsPerPixelY)
            Decoupe FrmPics.PicPions35, FrmPics.Grille35
        Case Else
            ' Au dessus de 768
            TailleJetons = 43
            FrmMain.Font.Size = 14
            FrmMain.StatusBar.Font.Size = 11
            Decoupe FrmPics.PicPions43, FrmPics.Grille43
    End Select
        
    ' Fixer la taille des jetons du joueur
    For i = 0 To 6
        PicLettreJoueur(i).Width = TailleJetons
        PicLettreJoueur(i).Height = TailleJetons
    Next
            
    ' Afficher les nombres horizontaux de 1 à 15
    For i = 1 To 15
        FrmMain.CurrentX = 32 + (TailleJetons + 1) * (i - 1) + (TailleJetons - TextWidth(CStr(i))) / 2
        FrmMain.CurrentY = 5
        FrmMain.Print CStr(i)
    Next

    ' Afficher les lettre verticales de A à O
    For i = 1 To 15
        FrmMain.CurrentX = 8
        FrmMain.CurrentY = 32 + (TailleJetons + 1) * (i - 1) + (TailleJetons - TextHeight(Chr(64 + i))) / 2
        FrmMain.Print Chr(64 + i)
    Next
    
    
    ' Afficher la grille
    PicGrille.Width = FrmPics.Grille.Width
    PicGrille.Height = FrmPics.Grille.Height
    PicGrille.PaintPicture FrmPics.Grille.Image, 0, 0, PicGrille.Width, PicGrille.Height
    
    ' Afficher la reglette
    PicReglette.Top = PicGrille.Top + PicGrille.Height + 17
    Select Case TailleJetons
        Case 35
            PicReglette.Height = FrmPics.PicReglette35.Height
            PicReglette.Width = FrmPics.PicReglette35.Width
            PicReglette.PaintPicture FrmPics.PicReglette35, 0, 0
        Case 43
            PicReglette.Height = FrmPics.PicReglette43.Height
            PicReglette.Width = FrmPics.PicReglette43.Width
            PicReglette.PaintPicture FrmPics.PicReglette43, 0, 0
    End Select
    PicReglette.Left = CInt(PicGrille.Left + (PicGrille.Width - PicReglette.Width) / 2)
    
    ' Placer les Frames du coté
    FrameOrthographe.Left = PicGrille.Left + PicGrille.Width + 20
    FrameHistorique.Left = PicGrille.Left + PicGrille.Width + 20
    FrameScores.Left = PicGrille.Left + PicGrille.Width + 20
    FrameChronometre.Left = PicGrille.Left + PicGrille.Width + 20
    
    ' Largeur de la fenêtre
    Me.Width = (FrmMain.FrameOrthographe.Left + FrmMain.FrameOrthographe.Width + 20) * Screen.TwipsPerPixelX
    FrmMain.StatusBar.Panels(1).Width = Me.Width / Screen.TwipsPerPixelX - 70
    
    ' Il faudrais calculer la hauteur de la barre des tâches...
    Me.Top = ((Screen.Height - (38 * Screen.TwipsPerPixelY)) - Me.Height) / 2
    Me.Left = (Screen.Width - Me.Width) / 2
    
    ' Positionner les boutons
    CmdJouer.Top = PicReglette.Top
    CmdJouer.Left = PicReglette.Left + PicReglette.Width + 20
    
    CmdPasser.Top = CmdJouer.Top + CmdJouer.Height + 4
    CmdPasser.Left = CmdJouer.Left
        
    CmdRAZJetons.Top = PicReglette.Top
    CmdRAZJetons.Left = CmdJouer.Left + CmdJouer.Width + 10
    
    CmdMelanger.Top = CmdJouer.Top + CmdJouer.Height + 4
    CmdMelanger.Left = CmdRAZJetons.Left
    
    CmdChanger.Top = CmdJouer.Top
    CmdChanger.Left = CmdMelanger.Left + CmdMelanger.Width + 10
    
    ' Les frames
    FrameOrthographe.Top = PicGrille.Top + PicGrille.Height - FrameOrthographe.Height
    
    
End Sub

Private Sub Decoupe(ImgSrc As PictureBox, ImgGrille As PictureBox)

    Dim i As Integer
    Dim j As Integer
    
    For i = 0 To 12
        FrmPics.PicLettre(i).PaintPicture ImgSrc.Image, 0, 0, TailleJetons, TailleJetons, i * TailleJetons, 0, TailleJetons, TailleJetons
        FrmPics.PicLettre(i).Refresh
        FrmPics.PicLettre(i).Width = TailleJetons
        FrmPics.PicLettre(i).Height = TailleJetons
        FrmPics.PicLettre(i + 13).PaintPicture ImgSrc.Image, 0, 0, TailleJetons, TailleJetons, i * TailleJetons, TailleJetons, TailleJetons, TailleJetons
        FrmPics.PicLettre(i + 13).Width = TailleJetons
        FrmPics.PicLettre(i + 13).Height = TailleJetons
    Next
    
    FrmPics.PicLettre(26).PaintPicture ImgSrc.Image, 0, 0, TailleJetons, TailleJetons, 13 * TailleJetons, TailleJetons, TailleJetons, TailleJetons
    FrmPics.PicLettre(26).Width = TailleJetons
    FrmPics.PicLettre(26).Height = TailleJetons

    FrmPics.PicTemp.Width = TailleJetons
    FrmPics.PicTemp.Height = TailleJetons
    
    FrmPics.Grille.Width = ImgGrille.Width
    FrmPics.Grille.Height = ImgGrille.Height
    FrmPics.Grille.PaintPicture ImgGrille.Image, 0, 0
End Sub

Private Sub AjouteHistorique(StrJoueur As String, MotScribble As ClsMotScribble)

    Dim Top As Integer
    Dim i As Integer

    If GridHistorique.RowHeight(1) = 0 Then
        GridHistorique.RowHeight(1) = GridHistorique.RowHeight(0)
        Top = 1
    Else
        GridHistorique.Rows = GridHistorique.Rows + 1
        Top = GridHistorique.Rows - 1
    End If
    
    GridHistorique.Row = Top
    
    GridHistorique.col = 0
    GridHistorique.CellAlignment = flexAlignLeftCenter
    GridHistorique.Text = CStr(Top)
    
    GridHistorique.col = 1
    GridHistorique.CellAlignment = flexAlignLeftCenter
    GridHistorique.Text = StrJoueur
    
    GridHistorique.col = 2
    GridHistorique.CellAlignment = flexAlignLeftCenter
    GridHistorique.Text = MotScribble.Mot
    
    GridHistorique.col = 3
    GridHistorique.CellAlignment = flexAlignCenterCenter
    GridHistorique.Text = MotScribble.Pos
    
    GridHistorique.col = 4
    GridHistorique.CellAlignment = flexAlignCenterCenter
    GridHistorique.Text = MotScribble.Pts
    
    GridHistorique.TopRow = Top
    
    For i = 0 To 4
        GridHistorique.col = i
        GridHistorique.CellBackColor = &HE0987B
    Next

    Top = Top - 1
    If Top > 0 Then
        GridHistorique.Row = Top
        For i = 0 To 4
            GridHistorique.col = i
            GridHistorique.CellBackColor = &HC0FFFF
        Next
    End If
End Sub

Private Sub TxtMotAVerifier_KeyPress(KeyAscii As Integer)

    Dim Car As String
        
    ' Filtrer les caractères frappés.
    Car = UCase(Chr(KeyAscii))
    If (Asc(Car) < Asc("A") Or Asc(Car) > Asc("Z")) Then
        If KeyAscii >= 32 Then
            KeyAscii = 0
        End If
    End If
        
    ' Vérifer la longeur de la zone de texte
    If KeyAscii >= 32 Then
        If Len(TxtMotAVerifier) > 14 Then KeyAscii = 0
    End If
    
End Sub

Private Sub TxtPts_GotFocus()

    GridHistorique.SetFocus
    
End Sub

Private Sub TxtPtsJoueur_GotFocus()

    GridHistorique.SetFocus
    
End Sub

Private Sub DuplicateTourSuivant()

    Dim i As Integer
    Dim MsPC As New ClsMotScribble
    Dim StrTemp As String
    
    Set MsPC = GDuplicate.GetSolution(NIVEAU)
        
    For i = 0 To 6
        PicLettreJoueur(i).Enabled = False
    Next
    
    DrawGrille
    ' Si le joueur a cliquer sur Passer
    If MsDuplicate.Pts = -1 Then MsDuplicate.Pts = 0
    
    If MsPC.Pts = 0 And MsDuplicate.Pts = 0 Then
        ' Fin de partie si aucune solution...
        FinSurJePasse
        Exit Sub
    End If
    
    NbJePasse = 0
    
    If MsPC.Pts > MsDuplicate.Pts Then
        AfficheMotEnBleu MsPC
        EnleverLettre MsPC, G
        G.PlacerMot MsPC
        GDuplicate.PlacerMot MsPC
        AjouteHistorique "PC", MsPC
    Else
        AfficheMotEnBleu MsDuplicate
        EnleverLettre MsDuplicate, G
        G.PlacerMot MsDuplicate
        GDuplicate.PlacerMot MsDuplicate
        AjouteHistorique "Joueur", MsDuplicate
    End If
    
    TxtPts = CStr(Val(TxtPts) + MsPC.Pts)
    TxtPtsJoueur = CStr(Val(TxtPtsJoueur) + MsDuplicate.Pts)
        
    MsDuplicate.Mot = ""
    MsDuplicate.Pos = ""
    MsDuplicate.Pts = 0
    
    G.IsPremierMot = False
    GDuplicate.IsPremierMot = False
    
        
    If DuplicateAjouteLettres() = False Then
        ' Fin de partie
        TimChrono.Enabled = False
    
        For i = 0 To 6
            PicLettreJoueur(i).Enabled = False
        Next
        CmdPasser.Enabled = False
        CmdMelanger.Enabled = False
        
        Load FrmFin
        StrTemp = "Joueur|" & TxtPtsJoueur & "|0|" & CStr(Val(TxtPtsJoueur))
        StrTemp = StrTemp & "|PC|" & TxtPts & "|0|" & CStr(Val(TxtPts))
        FrmFin.Affiche StrTemp
        FrmFin.Show vbModal
        
        MnuArreterLaPartie.Visible = False
        MnuNouvellePartie.Visible = True
        MnuParamertes.Enabled = True
        MnuChargerPartie.Enabled = True
        LblChrono = "00:00"
        Exit Sub
        
    End If
    
    If PAUSEAPRESCOUP = "oui" Then PauseApresCoupPc
        
    InitChrono
    GDuplicate.TrouveSolution (LETTRESUC)

End Sub

Private Function DuplicateAjouteLettres() As Boolean

    Dim i As Integer
    Dim j As Integer
    Dim Pos As Integer
    Dim Car As String
    Dim Pas As Integer
    Dim BackLettresUC As String
    Dim BackStrPioche As String
    Dim NoCoup As Integer
    
    Randomize (Time)
    Pas = (PicReglette.Width - 3) / 13
    
    For j = 0 To 12
        TabReglette(j).Lettre = ""
        TabReglette(j).IndexImg = 0
    Next
    
    For j = 0 To 6
        PicLettreJoueur(j).Tag = ""
        PicLettreJoueur(j).Visible = False
    Next
        
    ' Par défaut
    DuplicateAjouteLettres = True
        
    BackLettresUC = LETTRESUC
    BackStrPioche = STRPIOCHE
    NoCoup = Val(GridHistorique.TextMatrix(GridHistorique.Rows - 1, 0)) + 1
        
    Do
        LETTRESUC = BackLettresUC
        STRPIOCHE = BackStrPioche
        
        If Len(LETTRESUC) < 7 Then
            For i = 0 To 7
                Pos = Int(Rnd() * Len(STRPIOCHE)) + 1
                Car = Mid(STRPIOCHE, Pos, 1)
                If STRPIOCHE <> "" Then STRPIOCHE = Left(STRPIOCHE, Pos - 1) & Right(STRPIOCHE, Len(STRPIOCHE) - Pos)
                LETTRESUC = LETTRESUC + Car
                
                If STRPIOCHE = "" Then Exit For
                If Len(LETTRESUC) = 7 Then Exit For
            Next
        End If
        
        If NoCoup <= 15 Then
            ' Minimum 2 voyelles / 2 consomnes
            If CompteConsonnes(STRPIOCHE, True) >= 2 And CompteVoyelles(STRPIOCHE, True) >= 2 Then
                If CompteConsonnes(LETTRESUC, True) >= 2 And CompteVoyelles(LETTRESUC, True) >= 2 Then Exit Do
            Else
                ' on Joue s'il reste un Y ou ?
                If InStr(LETTRESUC, "Y") > 0 Or InStr(LETTRESUC, "?") > 0 Then Exit Do
                
                If Len(LETTRESUC) = 1 Or CompteConsonnes(LETTRESUC, True) = Len(LETTRESUC) Or CompteVoyelles(LETTRESUC, True) = Len(LETTRESUC) Then
                    ' Plus qu'une seule lettre ou que des consonnes ou que des voyelles
                    DuplicateAjouteLettres = False
                    'Exit Function
                End If
                ' On joue les lettres
                Exit Do
            End If
        End If
        
        If NoCoup >= 16 Then
            ' Minimum 1 voyelle /1 consomne
            If CompteConsonnes(STRPIOCHE, True) >= 1 And CompteVoyelles(STRPIOCHE, True) >= 1 Then
                If CompteConsonnes(LETTRESUC, True) >= 1 And CompteVoyelles(LETTRESUC, True) >= 1 Then Exit Do
            Else
                ' on Joue s'il reste un Y ou ?
                If InStr(LETTRESUC, "Y") > 0 Or InStr(LETTRESUC, "?") > 0 Then Exit Do
                
                If Len(LETTRESUC) = 1 Or CompteConsonnes(LETTRESUC, True) = Len(LETTRESUC) Or CompteVoyelles(LETTRESUC, True) = Len(LETTRESUC) Then
                    ' Plus qu'une seule lettre ou que des consonnes ou que des voyelles
                    DuplicateAjouteLettres = False
                    'Exit Function
                End If
                ' On joue les lettres
                Exit Do
            End If
        End If
        
    Loop While 1
        
    For i = 0 To Len(LETTRESUC) - 1
        Car = Mid(LETTRESUC, i + 1, 1)
        If Car <> "?" Then
            PicLettreJoueur(i).PaintPicture FrmPics.PicLettre(Asc(Car) - 65).Image, 0, 0
        Else
            PicLettreJoueur(i).PaintPicture FrmPics.PicLettre(26).Image, 0, 0
        End If
        
        PicLettreJoueur(i).Tag = LCase(Car)
        PicLettreJoueur(i).Visible = True
        PicLettreJoueur(i).Enabled = True
        PicLettreJoueur(i).Left = PicReglette.Left + (i + 3) * Pas + 2
        PicLettreJoueur(i).Top = PicReglette.Top + 2
        TabReglette(i + 3).Lettre = LCase(Car)
        TabReglette(i + 3).IndexImg = i
    Next
    
    Me.Font.Size = StatusBar.Font.Size
    If Me.TextWidth(STRPIOCHE) >= StatusBar.Panels(1).Width Then
        StatusBar.Panels(1).Text = CStr(Len(STRPIOCHE)) & " lettres."
    Else
        StatusBar.Panels(1).Text = STRPIOCHE
    End If
    
    StatusBar.Panels(2).Text = ""
    If Len(LETTRESUC) > 1 Then
        CmdMelanger.Enabled = True
        CmdPasser.Enabled = True
    End If

End Function

Private Function CompteConsonnes(StrToTest As String, CompteJocker As Boolean) As Integer

    Dim TabConsonnes As Variant
    Dim i As Integer
    Dim j As Integer
    
    TabConsonnes = Array("B", "C", "D", "F", "G", "H", "J", "K", "L", "M", "N", "P", "Q", "R", "S", "T", "V", "W", "X", "Z")
    If CompteJocker = True Then
        ReDim Preserve TabConsonnes(UBound(TabConsonnes) + 1)
        TabConsonnes(UBound(TabConsonnes)) = "?"
    End If
    
    For i = 1 To Len(StrToTest)
        For j = 0 To UBound(TabConsonnes)
            If Mid(StrToTest, i, 1) = TabConsonnes(j) Then CompteConsonnes = CompteConsonnes + 1
        Next
    Next
    
End Function

Private Function CompteVoyelles(StrToTest As String, CompteJocker As Boolean) As Integer

    Dim TabVoyelles As Variant
    Dim i As Integer
    Dim j As Integer
    
    TabVoyelles = Array("A", "E", "I", "O", "U", "Y")
    If CompteJocker = True Then
        ReDim Preserve TabVoyelles(UBound(TabVoyelles) + 1)
        TabVoyelles(UBound(TabVoyelles)) = "?"
    End If
    
    For i = 1 To Len(StrToTest)
        For j = 0 To UBound(TabVoyelles)
            If Mid(StrToTest, i, 1) = TabVoyelles(j) Then CompteVoyelles = CompteVoyelles + 1
        Next
    Next
    
End Function

Private Sub InitChrono()
    
    LblChrono.ForeColor = vbBlack
    TimChrono.Tag = CStr(GetTickCount())
    TimChrono_Timer
    DoEvents
    CmdFinTour.Enabled = False
    TimChrono.Enabled = True
    
End Sub
'
' PauseApresCoupPc :
' Faire clignoter le LblChrono toutes les 400 ms pendant TEMPSPAUSE secondes.
'
Private Sub PauseApresCoupPc()
    
    Dim T0 As Long
    Dim T As Long
        
    LblChrono.ForeColor = &HC0C0C0   ' &HA36C38
    LblChrono.BackColor = vbWhite
    
    T0 = GetTickCount() + (TEMPSPAUSE * 1000)
    
    Do
        LblChrono = ""
        T = GetTickCount + 500
        Do
            DoEvents
            If MnuArreterLaPartie.Visible = False Then Exit Sub
        Loop While GetTickCount() < T
        
        If PARTIECHRONOMETREE = "oui" Then
            LblChrono = Format(Int(CHRONOMETRE / 60), "00") & ":" & Format(CHRONOMETRE Mod 60, "00")
        Else
            LblChrono = "00:00"
        End If
        
        T = GetTickCount + 500
        Do
            DoEvents
            If MnuArreterLaPartie.Visible = False Then Exit Sub
        Loop While GetTickCount() < T
        
    Loop While GetTickCount < T0
    
    LblChrono.BackColor = &H80FFFF

End Sub

Private Sub FinSurJePasse()

    Dim i As Integer
    Dim Penalites As Integer
    Dim StrTemp As String
    
    Penalites = 0
    For i = 0 To 6
        PicLettreJoueur(i).Enabled = False
        If PicLettreJoueur(i).Tag <> "" Then
            Penalites = Penalites + G.PointLettres(PicLettreJoueur(i).Tag)
        End If
    Next
    StrTemp = "Joueur|" & TxtPtsJoueur & "|- " & CStr(Penalites) & "|" & CStr(Val(TxtPtsJoueur) - Penalites)
    TxtPtsJoueur = CStr(Val(TxtPtsJoueur) - Penalites)
    
    Penalites = 0
    For i = 1 To Len(LETTRESUC)
        Penalites = Penalites + G.PointLettres(LCase(Mid(LETTRESUC, i, 1)))
    Next
    StrTemp = StrTemp & "|PC|" & TxtPts & "|- " & CStr(Penalites) & "|" & CStr(Val(TxtPts) - Penalites)
    TxtPts = CStr(Val(TxtPts) - Penalites)
    
    Load FrmFin
    FrmFin.Affiche StrTemp
    FrmFin.Show vbModal
    
    MnuArreterLaPartie_Click

End Sub

Private Sub ChangeNLettres()

    Dim Diff As Integer
    Dim Car As String
    Dim j As Integer
    Dim i As Integer
    Dim Pos As Integer
    Dim StrLettresAChanger As String
    
    StrLettresAChanger = ""
    
    If CompteConsonnes(LETTRESUC, False) > CompteVoyelles(LETTRESUC, False) Then
        ' Changer Diff consonnes...
        Diff = Int((CompteConsonnes(LETTRESUC, False) - CompteVoyelles(LETTRESUC, False)) / 2) + 1
        i = 0
        Do
            j = Int(Rnd() * 7) + 1
            Car = Mid(LETTRESUC, j, 1)
            If CompteConsonnes(Car, False) = 1 Then
                StrLettresAChanger = StrLettresAChanger & Car
                LETTRESUC = Left(LETTRESUC, j - 1) & Right(LETTRESUC, Len(LETTRESUC) - j)
                i = i + 1
                If i = Diff Then Exit Do
            End If
        Loop While 1
    Else
        Diff = Int((CompteVoyelles(LETTRESUC, False) - CompteConsonnes(LETTRESUC, False)) / 2) + 1
        i = 0
        Do
            j = Int(Rnd() * 7) + 1
            Car = Mid(LETTRESUC, j, 1)
            If CompteVoyelles(Car, False) = 1 Then
                StrLettresAChanger = StrLettresAChanger & Car
                LETTRESUC = Left(LETTRESUC, j - 1) & Right(LETTRESUC, Len(LETTRESUC) - j)
                i = i + 1
                If i = Diff Then Exit Do
            End If
        Loop While 1
    End If
    
    ' Rajouter les lettres à changer à la pioche
    For i = 1 To Diff
        If Mid(StrLettresAChanger, i, 1) = "?" Then
            STRPIOCHE = STRPIOCHE & "?"
        Else
            For j = 1 To Len(STRPIOCHE)
                If Asc(Mid(STRPIOCHE, j, 1)) > Asc(Mid(StrLettresAChanger, i, 1)) Or Mid(STRPIOCHE, j, 1) = "?" Then
                    STRPIOCHE = Left(STRPIOCHE, j - 1) & Mid(StrLettresAChanger, i, 1) & Right(STRPIOCHE, Len(STRPIOCHE) - j + 1)
                    Exit For
                End If
            Next
        End If
    Next
        
    ' et rajouter des lettres à l'orinateur
    For i = 1 To Diff
        Pos = Int(Rnd() * Len(STRPIOCHE)) + 1
        Car = Mid(STRPIOCHE, Pos, 1)
        STRPIOCHE = Left(STRPIOCHE, Pos - 1) & Right(STRPIOCHE, Len(STRPIOCHE) - Pos)
        LETTRESUC = LETTRESUC + Car
    Next
    

    Load FrmMsgbox
    If Diff > 1 Then
        FrmMsgbox.LblMsg = "L'ordinateur change " & CStr(Diff) & " lettres !"
    Else
        FrmMsgbox.LblMsg = "L'ordinateur change 1 lettre !"
    End If
    FrmMsgbox.Show vbModal

End Sub
