VERSION 5.00
Begin VB.Form FrmDictionnaire 
   BackColor       =   &H00A36C38&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "  Changement de dictionnaire"
   ClientHeight    =   4680
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5520
   Icon            =   "FrmDictionnaire.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4680
   ScaleWidth      =   5520
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.Frame Frame1 
      BackColor       =   &H00A36C38&
      Caption         =   "Sélectionner le niveau de l'ordinateur"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2415
      Left            =   120
      TabIndex        =   1
      Top             =   120
      Width           =   5295
      Begin VB.OptionButton OptDico 
         BackColor       =   &H00A36C38&
         Caption         =   "ODS 5 - Championnats Francophones"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   0
         Left            =   240
         TabIndex        =   5
         Top             =   480
         Value           =   -1  'True
         Width           =   4815
      End
      Begin VB.OptionButton OptDico 
         BackColor       =   &H00A36C38&
         Caption         =   "TWL6 - Nord Américain"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   1
         Left            =   240
         TabIndex        =   4
         Top             =   960
         Width           =   2655
      End
      Begin VB.OptionButton OptDico 
         BackColor       =   &H00A36C38&
         Caption         =   "SOWPODS - Anglophone Championnats du monde"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   2
         Left            =   240
         TabIndex        =   3
         Top             =   1440
         Width           =   4935
      End
      Begin VB.OptionButton OptDico 
         BackColor       =   &H00A36C38&
         Caption         =   "ZINGA - Italien"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   3
         Left            =   240
         TabIndex        =   2
         Top             =   1920
         Width           =   2655
      End
   End
   Begin VB.CommandButton CmdOK 
      Caption         =   "&OK"
      Default         =   -1  'True
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   1320
      TabIndex        =   0
      Top             =   4200
      Width           =   2775
   End
   Begin VB.Label LblInfo 
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Dictionnaire italien basé sur le dictionnaire Zingarelli 2005"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1335
      Left            =   120
      TabIndex        =   6
      Top             =   2760
      Width           =   5295
   End
End
Attribute VB_Name = "FrmDictionnaire"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub CmdOK_Click()
    
    Dim StrNewDico As String
    
    If OptDico(0).Value = True Then StrNewDico = "ODS"
    If OptDico(1).Value = True Then StrNewDico = "TWL"
    If OptDico(2).Value = True Then StrNewDico = "SOWPODS"
    If OptDico(3).Value = True Then StrNewDico = "ZINGA"
   
    RegSaveString HKEY_LOCAL_MACHINE, KEYSCRIBBLE, "Dico", StrNewDico
    
    Unload Me
    
End Sub

Private Sub Form_Load()

    Dim StrDico As String
    
    StrDico = RegGetString(HKEY_LOCAL_MACHINE, KEYSCRIBBLE, "Dico")
    Select Case StrDico
        Case "ODS", ""
            OptDico(0).Value = True
            ShowInfo 0
        Case "TWL"
            OptDico(1).Value = True
            ShowInfo 1
        Case "SOWPODS"
            OptDico(2).Value = True
            ShowInfo 2
        Case "ZINGA"
            OptDico(3).Value = True
            ShowInfo 3
    End Select
    
End Sub

Private Sub OptDico_Click(Index As Integer)

    ShowInfo Index

End Sub

Private Sub ShowInfo(IntIndex As Integer)

    Select Case IntIndex
    
        Case 0
            LblInfo.Caption = "ODS5 : 378 989 mots - L’Officiel du Scrabble, 5ème édition (Larousse 2008) est utilisé dans les tous les clubs et championnats francophones."
        Case 1
            LblInfo.Caption = "TWL06 : 178 690 mots - Dictionnaire américain, basé sur l'Official Scrabble Players’ Dictionary, utilisé dans les tournois en Amérique du Nord."
        Case 2
            LblInfo.Caption = "SOWPODS : 216 553 mots - Combinaison des dictionnaires américain (OSPD) et anglais (Official Scrabble Words). Il est utilisé lors des Championnats du Monde de Scrabble anglophone, et dans les tournois organisés hors Amérique du Nord."
        Case 3
            LblInfo.Caption = "ZINGA : 584 983 mots - dictionnaire italien basé sur le dictionnaire Zingarelli 2005."
    End Select
    
End Sub
