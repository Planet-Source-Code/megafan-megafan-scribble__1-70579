VERSION 5.00
Begin VB.Form FrmMsgbox 
   BackColor       =   &H00A36C38&
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   1800
   ClientLeft      =   45
   ClientTop       =   45
   ClientWidth     =   5265
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   120
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   351
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.Timer Timer 
      Interval        =   100
      Left            =   120
      Top             =   1200
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00A36C38&
      Height          =   975
      Left            =   120
      TabIndex        =   1
      Top             =   120
      Width           =   5055
      Begin VB.Label LblMsg 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "L'ordinateur change x lettres !"
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
         Left            =   0
         TabIndex        =   2
         Top             =   360
         Width           =   5055
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
      Top             =   1320
      Width           =   2775
   End
End
Attribute VB_Name = "FrmMsgbox"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub CmdOK_Click()

    Unload Me
    
End Sub

Private Sub Timer_Timer()
    
    If PARTIECHRONOMETREE = "non" Then Exit Sub
    If FrmMain.LblChrono.Caption = "00:00" Then Unload Me
    
End Sub
