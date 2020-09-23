VERSION 5.00
Begin VB.Form FrmAProposDe 
   BackColor       =   &H00A36C38&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "  A propos de Megafan Scribble"
   ClientHeight    =   3675
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   6525
   Icon            =   "FrmAProposDe.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   245
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   435
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.Frame Frame2 
      BackColor       =   &H00A36C38&
      Caption         =   "Contact"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1215
      Left            =   240
      TabIndex        =   3
      Top             =   1920
      Width           =   5895
      Begin VB.Label Label6 
         BackStyle       =   0  'Transparent
         Caption         =   "http://www.vbfrance.com"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0080FFFF&
         Height          =   255
         Left            =   1920
         TabIndex        =   7
         Top             =   720
         Width           =   2295
      End
      Begin VB.Label Label5 
         BackStyle       =   0  'Transparent
         Caption         =   "megafan2001@yahoo.fr"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0080FFFF&
         Height          =   255
         Left            =   1920
         TabIndex        =   6
         Top             =   360
         Width           =   2295
      End
      Begin VB.Label Label4 
         BackStyle       =   0  'Transparent
         Caption         =   "Site Web :"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   600
         TabIndex        =   5
         Top             =   720
         Width           =   975
      End
      Begin VB.Label Label3 
         BackStyle       =   0  'Transparent
         Caption         =   "Par Mail : "
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   600
         TabIndex        =   4
         Top             =   360
         Width           =   975
      End
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
      Left            =   1920
      TabIndex        =   2
      Top             =   3240
      Width           =   2775
   End
   Begin VB.PictureBox Picture2 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   645
      Left            =   900
      ScaleHeight     =   645
      ScaleWidth      =   5280
      TabIndex        =   1
      Top             =   1080
      Width           =   5280
   End
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   645
      Left            =   240
      ScaleHeight     =   645
      ScaleWidth      =   3300
      TabIndex        =   0
      Top             =   240
      Width           =   3300
   End
   Begin VB.PictureBox Picture3 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   645
      Left            =   300
      ScaleHeight     =   645
      ScaleWidth      =   3300
      TabIndex        =   8
      Top             =   300
      Width           =   3300
   End
   Begin VB.PictureBox Picture4 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   645
      Left            =   900
      ScaleHeight     =   645
      ScaleWidth      =   5340
      TabIndex        =   9
      Top             =   1140
      Width           =   5340
   End
End
Attribute VB_Name = "FrmAProposDe"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hwnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long

Private Sub CmdFermer_Click()

    Unload Me
    
End Sub

Private Sub Form_Load()

    Picture1.Picture = FrmSplash.Picture1.Picture
    Picture2.Picture = FrmSplash.Picture2.Picture
    
    Label5.MousePointer = vbCustom
    Label5.MouseIcon = LoadResPicture(101, vbResCursor)
    
    Label6.MousePointer = vbCustom
    Label6.MouseIcon = LoadResPicture(101, vbResCursor)
    

End Sub

Private Sub Label5_Click()
    
    On Error Resume Next
    ShellExecute Me.hwnd, vbNullString, "mailto:megafan2001@yahoo.fr", vbNullString, "", 1

End Sub

Private Sub Label6_Click()

    On Error Resume Next
    ShellExecute Me.hwnd, vbNullString, "http://www.planet-source-code.com/", vbNullString, "", 1

End Sub
