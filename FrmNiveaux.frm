VERSION 5.00
Begin VB.Form FrmNiveaux 
   BackColor       =   &H00A36C38&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "  Niveaux de difficultés"
   ClientHeight    =   3720
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   4215
   Icon            =   "FrmNiveaux.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   248
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   281
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
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
      Left            =   600
      TabIndex        =   6
      Top             =   3240
      Width           =   2775
   End
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
      Height          =   3015
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   3975
      Begin VB.OptionButton OptNiveau 
         BackColor       =   &H00A36C38&
         Caption         =   "Champion (meilleurs coups)"
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
         Index           =   4
         Left            =   600
         TabIndex        =   5
         Top             =   2400
         Width           =   3255
      End
      Begin VB.OptionButton OptNiveau 
         BackColor       =   &H00A36C38&
         Caption         =   "Très bon joueur"
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
         Left            =   600
         TabIndex        =   4
         Top             =   1920
         Width           =   2655
      End
      Begin VB.OptionButton OptNiveau 
         BackColor       =   &H00A36C38&
         Caption         =   "Moyen"
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
         Left            =   600
         TabIndex        =   3
         Top             =   1440
         Value           =   -1  'True
         Width           =   2655
      End
      Begin VB.OptionButton OptNiveau 
         BackColor       =   &H00A36C38&
         Caption         =   "Débutant"
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
         Left            =   600
         TabIndex        =   2
         Top             =   960
         Width           =   2655
      End
      Begin VB.OptionButton OptNiveau 
         BackColor       =   &H00A36C38&
         Caption         =   "Trés facile (enfants)"
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
         Left            =   600
         TabIndex        =   1
         Top             =   480
         Width           =   2655
      End
   End
End
Attribute VB_Name = "FrmNiveaux"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub CmdOK_Click()

    Dim n As Long
    
    For n = 0 To 4
        If OptNiveau(n).Value = True Then Exit For
    Next
    
    NIVEAU = n + 1
    RegSaveDword HKEY_LOCAL_MACHINE, KEYSCRIBBLE, "Niveau", n + 1
    Unload Me

End Sub

Private Sub Form_Load()
    
    Dim n As Long

    n = RegGetDword(HKEY_LOCAL_MACHINE, KEYSCRIBBLE, "Niveau")
    If n <> 0 Then OptNiveau(n - 1).Value = True
    
End Sub
