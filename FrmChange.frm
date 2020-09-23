VERSION 5.00
Begin VB.Form FrmChange 
   BackColor       =   &H00A36C38&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "  Echange de jetons"
   ClientHeight    =   2550
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7080
   Icon            =   "FrmChange.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   170
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   472
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.Timer Timer 
      Interval        =   100
      Left            =   120
      Top             =   2040
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00A36C38&
      Caption         =   "Sélectionner les lettres à changer"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1815
      Left            =   240
      TabIndex        =   1
      Top             =   120
      Width           =   6495
      Begin VB.CheckBox ChkJetons 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         ForeColor       =   &H80000008&
         Height          =   855
         Index           =   0
         Left            =   480
         MaskColor       =   &H00FFFFFF&
         Style           =   1  'Graphical
         TabIndex        =   8
         Top             =   600
         UseMaskColor    =   -1  'True
         Width           =   735
      End
      Begin VB.CheckBox ChkJetons 
         BackColor       =   &H00E0E0E0&
         Height          =   855
         Index           =   1
         Left            =   1320
         MaskColor       =   &H00E0E0E0&
         Style           =   1  'Graphical
         TabIndex        =   7
         Top             =   600
         Width           =   735
      End
      Begin VB.CheckBox ChkJetons 
         BackColor       =   &H00E0E0E0&
         Height          =   855
         Index           =   2
         Left            =   2160
         MaskColor       =   &H00E0E0E0&
         Style           =   1  'Graphical
         TabIndex        =   6
         Top             =   600
         Width           =   735
      End
      Begin VB.CheckBox ChkJetons 
         BackColor       =   &H00E0E0E0&
         Height          =   855
         Index           =   3
         Left            =   3000
         MaskColor       =   &H00E0E0E0&
         Style           =   1  'Graphical
         TabIndex        =   5
         Top             =   600
         Width           =   735
      End
      Begin VB.CheckBox ChkJetons 
         BackColor       =   &H00E0E0E0&
         Height          =   855
         Index           =   4
         Left            =   3840
         MaskColor       =   &H00E0E0E0&
         Style           =   1  'Graphical
         TabIndex        =   4
         Top             =   600
         Width           =   735
      End
      Begin VB.CheckBox ChkJetons 
         BackColor       =   &H00E0E0E0&
         Height          =   855
         Index           =   5
         Left            =   4680
         MaskColor       =   &H00E0E0E0&
         Style           =   1  'Graphical
         TabIndex        =   3
         Top             =   600
         Width           =   735
      End
      Begin VB.CheckBox ChkJetons 
         BackColor       =   &H00E0E0E0&
         Height          =   855
         Index           =   6
         Left            =   5520
         MaskColor       =   &H00E0E0E0&
         Style           =   1  'Graphical
         TabIndex        =   2
         Top             =   600
         Width           =   735
      End
      Begin VB.PictureBox PicSel 
         BackColor       =   &H00E0987B&
         BorderStyle     =   0  'None
         Height          =   1095
         Index           =   0
         Left            =   360
         ScaleHeight     =   73
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   65
         TabIndex        =   9
         Top             =   480
         Visible         =   0   'False
         Width           =   975
      End
      Begin VB.PictureBox PicSel 
         BackColor       =   &H00E0987B&
         BorderStyle     =   0  'None
         Height          =   1095
         Index           =   1
         Left            =   1200
         ScaleHeight     =   73
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   65
         TabIndex        =   10
         Top             =   480
         Visible         =   0   'False
         Width           =   975
      End
      Begin VB.PictureBox PicSel 
         BackColor       =   &H00E0987B&
         BorderStyle     =   0  'None
         Height          =   1095
         Index           =   2
         Left            =   2160
         ScaleHeight     =   73
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   65
         TabIndex        =   11
         Top             =   480
         Visible         =   0   'False
         Width           =   975
      End
      Begin VB.PictureBox PicSel 
         BackColor       =   &H00E0987B&
         BorderStyle     =   0  'None
         Height          =   1095
         Index           =   3
         Left            =   3120
         ScaleHeight     =   73
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   65
         TabIndex        =   12
         Top             =   480
         Visible         =   0   'False
         Width           =   975
      End
      Begin VB.PictureBox PicSel 
         BackColor       =   &H00E0987B&
         BorderStyle     =   0  'None
         Height          =   1095
         Index           =   4
         Left            =   4080
         ScaleHeight     =   73
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   65
         TabIndex        =   13
         Top             =   480
         Visible         =   0   'False
         Width           =   975
      End
      Begin VB.PictureBox PicSel 
         BackColor       =   &H00E0987B&
         BorderStyle     =   0  'None
         Height          =   1095
         Index           =   5
         Left            =   5040
         ScaleHeight     =   73
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   65
         TabIndex        =   14
         Top             =   480
         Visible         =   0   'False
         Width           =   975
      End
      Begin VB.PictureBox PicSel 
         BackColor       =   &H00E0987B&
         BorderStyle     =   0  'None
         Height          =   1095
         Index           =   6
         Left            =   6000
         ScaleHeight     =   73
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   65
         TabIndex        =   15
         Top             =   480
         Visible         =   0   'False
         Width           =   975
      End
   End
   Begin VB.CommandButton CmdOK 
      Caption         =   "&OK"
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
      Left            =   2400
      TabIndex        =   0
      Top             =   2040
      Width           =   2250
   End
End
Attribute VB_Name = "FrmChange"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub ChkJetons_Click(Index As Integer)

    If ChkJetons(Index).Value = vbChecked Then
        PicSel(Index).Visible = True
    Else
        PicSel(Index).Visible = False
    End If

End Sub

Private Sub CmdOK_Click()

    Dim i As Integer
    Dim StrIndexJetons As String
    
    StrIndexJetons = ""
    For i = 0 To 6
        If ChkJetons(i).Value = vbChecked Then
            StrIndexJetons = StrIndexJetons & CStr(i) & ","
        End If
    Next
    
    Me.Tag = StrIndexJetons
    Me.Hide

End Sub

Private Sub Form_Load()
    
    Dim i As Integer
    Dim SX As Long
    Dim SY As Long
    
    SX = Screen.TwipsPerPixelX
    SY = Screen.TwipsPerPixelY
    
    For i = 0 To 6
        ChkJetons(i).Picture = FrmMain.PicLettreJoueur(i).Image
        ChkJetons(i).Width = FrmMain.PicLettreJoueur(i).Width * SX
        ChkJetons(i).Height = FrmMain.PicLettreJoueur(i).Height * SY
        If i > 0 Then
            ChkJetons(i).Left = ChkJetons(i - 1).Left + ChkJetons(i).Width + 10 * SX
        End If
        PicSel(i).Left = ChkJetons(i).Left - 4 * SX
        PicSel(i).Top = ChkJetons(i).Top - 4 * SY
        PicSel(i).Height = ChkJetons(i).Height + 8 * SX
        PicSel(i).Width = ChkJetons(i).Width + 8 * SY
    Next
    
End Sub

Private Sub Timer_Timer()

    If PARTIECHRONOMETREE = "non" Then Exit Sub
    If FrmMain.LblChrono.Caption = "00:00" Then Unload Me
    
End Sub
