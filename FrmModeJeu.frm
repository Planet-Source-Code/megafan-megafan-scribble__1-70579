VERSION 5.00
Begin VB.Form FrmModeDeJeu 
   BackColor       =   &H00A36C38&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "  Mode de jeu"
   ClientHeight    =   4560
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   5370
   Icon            =   "FrmModeJeu.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   304
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   358
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.Frame Frame2 
      BackColor       =   &H00A36C38&
      Caption         =   "Options"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1455
      Left            =   360
      TabIndex        =   5
      Top             =   2280
      Width           =   4695
      Begin VB.ComboBox CmbPause 
         BackColor       =   &H00C0FFFF&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   3360
         Style           =   2  'Dropdown List
         TabIndex        =   9
         Top             =   930
         Width           =   975
      End
      Begin VB.CheckBox ChkPause 
         BackColor       =   &H00A36C38&
         Caption         =   "Pause après le coup du PC"
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
         Left            =   480
         TabIndex        =   8
         Top             =   960
         Width           =   2775
      End
      Begin VB.ComboBox CmbChrono 
         BackColor       =   &H00C0FFFF&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   3360
         Style           =   2  'Dropdown List
         TabIndex        =   7
         Top             =   450
         Width           =   975
      End
      Begin VB.CheckBox ChkChrono 
         BackColor       =   &H00A36C38&
         Caption         =   "Partie chronomètrée"
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
         Left            =   480
         TabIndex        =   6
         Top             =   480
         Width           =   2295
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00A36C38&
      Caption         =   "Sélectionner le type de jeu"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1935
      Left            =   360
      TabIndex        =   1
      Top             =   240
      Width           =   4695
      Begin VB.OptionButton OptTypeJeu 
         BackColor       =   &H00A36C38&
         Caption         =   "Normal (Joueur contre PC)"
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
         Left            =   480
         TabIndex        =   4
         Top             =   480
         Value           =   -1  'True
         Width           =   2655
      End
      Begin VB.OptionButton OptTypeJeu 
         BackColor       =   &H00A36C38&
         Caption         =   "Duplicate (Même pioche)"
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
         Left            =   480
         TabIndex        =   3
         Top             =   960
         Width           =   2655
      End
      Begin VB.OptionButton OptTypeJeu 
         BackColor       =   &H00A36C38&
         Caption         =   "Solo (Faites le meilleur score)"
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
         Left            =   480
         TabIndex        =   2
         Top             =   1440
         Width           =   3135
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
      Left            =   1200
      TabIndex        =   0
      Top             =   4080
      Width           =   2775
   End
End
Attribute VB_Name = "FrmModeDeJeu"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub ChkChrono_Click()

    If ChkChrono.Value = vbChecked Then
        CmbChrono.Enabled = True
        ChkPause.Enabled = True
    Else
        CmbChrono.Enabled = False
        ChkPause.Enabled = False
        ChkPause.Value = vbUnchecked
    End If
    
End Sub

Private Sub ChkPause_Click()

    If ChkPause.Value = vbChecked Then
        CmbPause.Enabled = True
    Else
        CmbPause.Enabled = False
    End If

End Sub

Private Sub CmdOK_Click()

    Dim n As Long
    
    For n = 0 To 2
        If OptTypeJeu(n).Value = True Then Exit For
    Next
    
    MODEJEU = n
    RegSaveDword HKEY_LOCAL_MACHINE, KEYSCRIBBLE, "ModeJeu", n
    
    If ChkChrono.Value = vbChecked Then
        RegSaveString HKEY_LOCAL_MACHINE, KEYSCRIBBLE, "PartieChronometree", "oui"
        PARTIECHRONOMETREE = "oui"
    Else
        RegSaveString HKEY_LOCAL_MACHINE, KEYSCRIBBLE, "PartieChronometree", "non"
        PARTIECHRONOMETREE = "non"
    End If
    
    RegSaveDword HKEY_LOCAL_MACHINE, KEYSCRIBBLE, "Chrono", (CmbChrono.ListIndex + 1) * 30
    CHRONOMETRE = (CmbChrono.ListIndex + 1) * 30
    
    
    If ChkPause.Value = vbChecked Then
        RegSaveString HKEY_LOCAL_MACHINE, KEYSCRIBBLE, "PauseApresCoup", "oui"
        PAUSEAPRESCOUP = "oui"
    Else
        RegSaveString HKEY_LOCAL_MACHINE, KEYSCRIBBLE, "PauseApresCoup", "non"
        PAUSEAPRESCOUP = "non"
    End If
    
    RegSaveDword HKEY_LOCAL_MACHINE, KEYSCRIBBLE, "TempsPause", CmbPause.ListIndex + 1
    TEMPSPAUSE = CmbPause.ListIndex + 1
    
    
    Unload Me
    DoEvents

End Sub

Private Sub Form_Load()
    
    Dim n As Long

    n = RegGetDword(HKEY_LOCAL_MACHINE, KEYSCRIBBLE, "ModeJeu")
    If n < 0 Or n > 2 Then n = 0
    OptTypeJeu(n).Value = True
    
    If RegGetString(HKEY_LOCAL_MACHINE, KEYSCRIBBLE, "PartieChronometree") = "oui" Then
        ChkChrono.Value = vbChecked
        CmbChrono.Enabled = True
        CmbPause.Enabled = False
    Else
        ChkChrono.Value = vbUnchecked
        CmbChrono.Enabled = False
        CmbPause.Enabled = True
    End If
    
    For n = 1 To 20
        CmbChrono.AddItem Format(Int(n / 2), "00:") & Format(n * 30 Mod 60, "00")
        If RegGetDword(HKEY_LOCAL_MACHINE, KEYSCRIBBLE, "Chrono") = n * 30 Then
            CmbChrono.ListIndex = n - 1
        End If
    Next
    
    If CmbChrono.ListIndex = -1 Then CmbChrono.ListIndex = 5
    
    '//////////////////////////////////////////////////////////////////////////////////////////////
    
    If RegGetString(HKEY_LOCAL_MACHINE, KEYSCRIBBLE, "PauseApresCoup") = "oui" Then
        ChkPause.Value = vbChecked
        CmbPause.Enabled = True
    Else
        ChkPause.Value = vbUnchecked
        CmbPause.Enabled = False
    End If
    
    For n = 1 To 10
        CmbPause.AddItem CStr(n)
        If RegGetDword(HKEY_LOCAL_MACHINE, KEYSCRIBBLE, "TempsPause") = n Then
            CmbPause.ListIndex = n - 1
        End If
    Next
    
    If CmbPause.ListIndex = -1 Then CmbPause.ListIndex = 2
    
End Sub

Private Sub OptTypeJeu_Click(Index As Integer)

    If OptTypeJeu(2).Value = True Then
        ChkChrono.Enabled = False
        ChkChrono.Value = vbUnchecked
        ChkChrono_Click
        ChkPause.Enabled = False
    Else
        ChkChrono.Enabled = True
        ChkPause.Enabled = True
    End If

End Sub
