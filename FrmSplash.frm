VERSION 5.00
Begin VB.Form FrmSplash 
   BackColor       =   &H00A36C38&
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   2385
   ClientLeft      =   45
   ClientTop       =   45
   ClientWidth     =   6390
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   159
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   426
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.PictureBox Picture2 
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   660
      Left            =   885
      Picture         =   "FrmSplash.frx":0000
      ScaleHeight     =   660
      ScaleWidth      =   5280
      TabIndex        =   1
      Top             =   1080
      Width           =   5280
   End
   Begin VB.PictureBox Picture1 
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   660
      Left            =   240
      Picture         =   "FrmSplash.frx":B5C2
      ScaleHeight     =   660
      ScaleWidth      =   4620
      TabIndex        =   0
      Top             =   240
      Width           =   4620
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
      ScaleWidth      =   4620
      TabIndex        =   3
      Top             =   300
      Width           =   4620
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
      TabIndex        =   4
      Top             =   1140
      Width           =   5340
   End
   Begin VB.Label LblMessage 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Chargement en cours..."
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0080FFFF&
      Height          =   375
      Left            =   120
      TabIndex        =   2
      Top             =   1920
      Width           =   6135
   End
End
Attribute VB_Name = "FrmSplash"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
