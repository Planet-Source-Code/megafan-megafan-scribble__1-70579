VERSION 5.00
Begin VB.Form Frm1erTour 
   BackColor       =   &H00A36C38&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "  Définition du 1er tour"
   ClientHeight    =   2535
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   3960
   Icon            =   "Frm1erTour.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   169
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   264
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton CmdOk 
      Caption         =   "&OK"
      Default         =   -1  'True
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
      Left            =   960
      TabIndex        =   4
      Top             =   2040
      Width           =   2055
   End
   Begin VB.OptionButton Opt1erTour 
      BackColor       =   &H00A36C38&
      Caption         =   "Aléatoire"
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
      Index           =   0
      Left            =   600
      TabIndex        =   2
      Top             =   600
      Value           =   -1  'True
      Width           =   3135
   End
   Begin VB.OptionButton Opt1erTour 
      BackColor       =   &H00A36C38&
      Caption         =   "Ordinateur"
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
      Index           =   2
      Left            =   600
      TabIndex        =   1
      Top             =   1560
      Width           =   3135
   End
   Begin VB.OptionButton Opt1erTour 
      BackColor       =   &H00A36C38&
      Caption         =   "Joueur"
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
      Index           =   1
      Left            =   600
      TabIndex        =   0
      Top             =   1080
      Width           =   3135
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Chosir qui commence :"
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
      Left            =   600
      TabIndex        =   3
      Top             =   120
      Width           =   3255
   End
End
Attribute VB_Name = "Frm1erTour"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub CmdOK_Click()

    Dim p As Long
    
    For p = 0 To 2
        If Opt1erTour(p).Value = True Then Exit For
    Next
    
    PREMIERTOUR = p
    RegSaveDword HKEY_LOCAL_MACHINE, KEYSCRIBBLE, "PremierTour", p
    Unload Me
    
End Sub

Private Sub Form_Load()

    Dim p As Long

    p = RegGetDword(HKEY_LOCAL_MACHINE, KEYSCRIBBLE, "PremierTour")
    Opt1erTour(p).Value = True
    
End Sub
