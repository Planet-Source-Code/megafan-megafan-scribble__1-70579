VERSION 5.00
Begin VB.Form FrmJocker 
   BackColor       =   &H00A36C38&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "  Jocker"
   ClientHeight    =   1290
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   3195
   Icon            =   "FrmJocker.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   86
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   213
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.Timer Timer 
      Interval        =   100
      Left            =   240
      Top             =   840
   End
   Begin VB.TextBox TxtLettre 
      Alignment       =   2  'Center
      BackColor       =   &H00C0FFFF&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   2040
      TabIndex        =   0
      Top             =   180
      Width           =   615
   End
   Begin VB.CommandButton CmdOk 
      Caption         =   "&OK"
      Default         =   -1  'True
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
      Height          =   375
      Left            =   840
      TabIndex        =   1
      Top             =   840
      Width           =   1455
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Chosir une lettre "
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   240
      TabIndex        =   2
      Top             =   240
      Width           =   1815
   End
End
Attribute VB_Name = "FrmJocker"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub CmdOK_Click()
    
    Me.Hide

End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)

    If UnloadMode = 0 Then Cancel = 1

End Sub

Private Sub Timer_Timer()

    If PARTIECHRONOMETREE = "non" Then Exit Sub
    If FrmMain.LblChrono.Caption = "00:00" Then Unload Me

End Sub

Private Sub TxtLettre_KeyPress(KeyAscii As Integer)

    Dim Car As String

    CmdOK.Enabled = False
    TxtLettre = ""
    Car = UCase(Chr(KeyAscii))
    If Asc(Car) < Asc("A") Or Asc(Car) > Asc("Z") Then
        KeyAscii = 0
    Else
        If KeyAscii >= 97 Then KeyAscii = KeyAscii - 32
        CmdOK.Enabled = True
    End If
    
End Sub
