VERSION 5.00
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form FrmAnagrammes 
   BackColor       =   &H00A36C38&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "  Anagrammes"
   ClientHeight    =   6135
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   3270
   Icon            =   "FrmAnagrammes.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   409
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   218
   ShowInTaskbar   =   0   'False
   Begin MSComCtl2.FlatScrollBar Vscroll 
      Height          =   4845
      Left            =   2700
      TabIndex        =   5
      Top             =   510
      Width           =   285
      _ExtentX        =   503
      _ExtentY        =   8546
      _Version        =   393216
      Appearance      =   0
      Orientation     =   1245184
   End
   Begin VB.CommandButton CmdFermer 
      Caption         =   "Fermer"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Left            =   240
      TabIndex        =   4
      Top             =   5760
      Width           =   2775
   End
   Begin VB.CommandButton CmdChercher 
      Caption         =   "Chercher"
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
      Height          =   330
      Left            =   1920
      TabIndex        =   2
      Top             =   120
      Width           =   1095
   End
   Begin VB.TextBox txtAnagram 
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
      Height          =   330
      Left            =   240
      TabIndex        =   0
      Top             =   120
      Width           =   1575
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid Grid 
      Height          =   4890
      Left            =   240
      TabIndex        =   1
      Top             =   480
      Width           =   2775
      _ExtentX        =   4895
      _ExtentY        =   8625
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
   Begin VB.Label LblNbcomb 
      BackStyle       =   0  'Transparent
      Caption         =   "- mots en - s"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   240
      TabIndex        =   3
      Top             =   5400
      Width           =   2655
   End
End
Attribute VB_Name = "FrmAnagrammes"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Declare Function timeGetTime Lib "winmm.dll" () As Long
Private OldRow As Long
Private Ana2 As ClsMots

Private Sub CmdChercher_Click()
    
    Dim StartTime As Single
    Dim i As Long
    Dim FinalTime As Single
    Dim NbMots As Long
    Dim StrMot As String
    
    LblNbcomb = Format(Len(txtAnagram)) & " lettres"
    StartTime = timeGetTime

    CmdChercher.Enabled = False
    NbMots = Ana2.AnaGramme(txtAnagram.Text)
    
    If NbMots > 10000 Then
        CmdChercher.Enabled = True
        Exit Sub
    End If
         
    Ana2.ClasserMotsParLongueur
    Grid.Rows = 2
    Vscroll.Enabled = True
    For i = 1 To Ana2.TotalMots
        StrMot = Ana2.GetMot(i)
        If InStr(StrMot, " ") Then
            Grid.AddItem StrMot & vbTab & Len(Left(StrMot, InStr(StrMot, " ") - 1))
        Else
            Grid.AddItem StrMot & vbTab & Len(StrMot)
        End If
    Next
    
    FinalTime = (timeGetTime - StartTime) / 1000
    If Grid.Rows > 2 Then
        Grid.RemoveItem 1
        OldRow = 1
        Vscroll.Min = 1
        Vscroll.Max = NbMots
        LblNbcomb = Format(NbMots) & " mots en " & Format(FinalTime) & " secondes"
    Else
        Grid.RowHeight(1) = 0
        Vscroll.Enabled = False
        LblNbcomb = "0 mots en " & Format(FinalTime) & " secondes"
    End If
    
    CmdChercher.Enabled = True

End Sub

Private Sub CmdFermer_Click()

    Unload Me

End Sub

Private Sub Form_Load()
    
    Set Ana2 = New ClsMots
    Ana2.Dictionnaire = Dico
    
    Grid.FormatString = "^Mot|^Lg"
    Grid.ColWidth(0) = 1900
    Grid.ColWidth(1) = 500
    Grid.ColAlignment(0) = flexAlignLeftCenter
    Grid.ColAlignment(1) = flexAlignCenterCenter
    Grid.RowHeight(1) = 0
    
    Me.Left = FrmMain.Left + (FrmMain.PicGrille.Left + ((FrmMain.PicGrille.Width - Me.Width / Screen.TwipsPerPixelX) / 2)) * Screen.TwipsPerPixelX
    Me.Top = FrmMain.Top + (FrmMain.PicGrille.Top + FrmMain.PicGrille.Top + 60) * Screen.TwipsPerPixelY
    
End Sub

Private Sub Grid_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)

    Dim i As Long
    Dim TempRow As Long
    Dim StrMot As String

    If y < Grid.RowHeight(0) Then
    
        If Grid.RowHeight(1) = 0 Then Exit Sub
        
        If x < Grid.ColWidth(0) Then
            Ana2.ClasserMotParOrdreAlphabetique
        Else
            Ana2.ClasserMotsParLongueur
        End If
        
        Grid.Rows = 2
        
        For i = 1 To Ana2.TotalMots
            StrMot = Ana2.GetMot(i)
            If InStr(StrMot, " ") Then
                Grid.AddItem StrMot & vbTab & Len(Left(StrMot, InStr(StrMot, " ") - 1))
            Else
                Grid.AddItem StrMot & vbTab & Len(StrMot)
            End If
        Next
        
        If Grid.Rows > 2 Then
            Grid.RemoveItem 1
        Else
            Grid.RowHeight(1) = 0
        End If
        
        Vscroll.Value = 1
    
    Else
        
        If OldRow = 0 Then Exit Sub
        
        TempRow = Grid.Row
        
        Grid.Row = OldRow
        Grid.col = 0
        Grid.CellBackColor = &HC0FFFF
        Grid.col = 1
        Grid.CellBackColor = &HC0FFFF
        
        Grid.Row = TempRow
        Grid.col = 0
        Grid.CellBackColor = &HE0987B
        Grid.col = 1
        Grid.CellBackColor = &HE0987B
        
        OldRow = Grid.Row
        
    End If

End Sub

Private Sub txtAnagram_KeyPress(KeyAscii As Integer)

    Dim Car As String
    Dim i As Integer
    Dim Compte As Integer
    
    ' Filtrer les caractères frappés.
    Car = UCase(Chr(KeyAscii))
    If (Asc(Car) < Asc("A") Or Asc(Car) > Asc("Z")) And Asc(Car) <> Asc("?") Then
        If KeyAscii >= 32 Then
            KeyAscii = 0
        End If
    End If
    
    ' Compter le nombre de jockers
    If KeyAscii = 63 Then ' ?
        Compte = 0
        For i = 1 To Len(txtAnagram)
            If Mid(txtAnagram, i, 1) = "?" Then Compte = Compte + 1
        Next
        If Compte > 2 Then KeyAscii = 0  ' donc 4éme jocker frappé
    End If
    
    ' Vérifer la longeur de la zone de texte
    If KeyAscii >= 32 Then
        If Len(txtAnagram) > 14 Then KeyAscii = 0
    End If
    
End Sub

Private Sub Vscroll_Change()

    Grid.TopRow = Vscroll.Value

End Sub

Private Sub Vscroll_Scroll()

    Grid.TopRow = Vscroll.Value

End Sub
