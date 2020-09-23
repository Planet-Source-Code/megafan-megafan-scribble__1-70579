VERSION 5.00
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Begin VB.Form FrmFin 
   BackColor       =   &H00A36C38&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "  Fin de la partie"
   ClientHeight    =   2130
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   4365
   Icon            =   "FrmFin.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   142
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   291
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton CmdFermer 
      Caption         =   "&Fermer"
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
      Left            =   840
      TabIndex        =   1
      Top             =   1680
      Width           =   2895
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid GridFin 
      Height          =   855
      Left            =   120
      TabIndex        =   2
      Top             =   600
      Width           =   4125
      _ExtentX        =   7276
      _ExtentY        =   1508
      _Version        =   393216
      BackColor       =   12648447
      Rows            =   3
      Cols            =   4
      FixedCols       =   0
      BackColorBkg    =   8454143
      FocusRect       =   0
      HighLight       =   0
      ScrollBars      =   0
      BorderStyle     =   0
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
      _Band(0).Cols   =   4
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Partie terminée !"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   0
      TabIndex        =   0
      Top             =   120
      Width           =   4335
   End
End
Attribute VB_Name = "FrmFin"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub CmdFermer_Click()

    Unload Me
    
End Sub

Private Sub Form_Load()

    GridFin.FormatString = "^Nom|^Score|^Penalité|^Score fin"
    GridFin.ColWidth(0) = 1100
    GridFin.ColWidth(1) = 1000
    GridFin.ColWidth(2) = 1000
    GridFin.ColWidth(3) = 1100
    
End Sub

Public Sub Affiche(StrScores As String)

    Dim TabMots() As String
    Dim i As Integer
    Dim j As Integer
    TabMots = Split(StrScores, "|")
        
    GridFin.Row = 1
    j = 0
    
    For i = 0 To 7
        GridFin.col = j
        If j = 0 Then GridFin.CellAlignment = flexAlignLeftCenter
        GridFin.Text = TabMots(i)
        j = j + 1
        If i = 3 Then
            GridFin.Row = 2
            j = 0
        End If
    Next

End Sub
