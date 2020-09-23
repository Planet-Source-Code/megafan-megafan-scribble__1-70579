VERSION 5.00
Begin VB.Form FrmPics 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "  Collection d'images"
   ClientHeight    =   11010
   ClientLeft      =   150
   ClientTop       =   540
   ClientWidth     =   11475
   Icon            =   "FrmPics.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   734
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   765
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.PictureBox PicReglette35 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   585
      Left            =   600
      Picture         =   "FrmPics.frx":5F32
      ScaleHeight     =   39
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   471
      TabIndex        =   34
      Top             =   6840
      Width           =   7065
   End
   Begin VB.PictureBox PicReglette43 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   705
      Left            =   4680
      Picture         =   "FrmPics.frx":1372C
      ScaleHeight     =   47
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   575
      TabIndex        =   29
      Top             =   4680
      Width           =   8625
   End
   Begin VB.PictureBox PicPions35 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   1050
      Left            =   600
      Picture         =   "FrmPics.frx":274AE
      ScaleHeight     =   1050
      ScaleWidth      =   7350
      TabIndex        =   32
      Top             =   5520
      Width           =   7350
   End
   Begin VB.PictureBox PicPions43 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   1290
      Left            =   3120
      Picture         =   "FrmPics.frx":40770
      ScaleHeight     =   1290
      ScaleWidth      =   9030
      TabIndex        =   31
      Top             =   3120
      Width           =   9030
   End
   Begin VB.PictureBox Grille43 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   9930
      Left            =   3000
      Picture         =   "FrmPics.frx":66712
      ScaleHeight     =   662
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   662
      TabIndex        =   30
      Top             =   0
      Width           =   9930
   End
   Begin VB.PictureBox PicTemp 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "CG Omega"
         Size            =   20.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   645
      Left            =   9480
      ScaleHeight     =   43
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   43
      TabIndex        =   28
      Top             =   10440
      Width           =   645
   End
   Begin VB.PictureBox PicLettre 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "CG Omega"
         Size            =   20.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   645
      Index           =   1
      Left            =   840
      ScaleHeight     =   43
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   43
      TabIndex        =   27
      Top             =   10440
      Width           =   645
   End
   Begin VB.PictureBox PicLettre 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "CG Omega"
         Size            =   20.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   645
      Index           =   2
      Left            =   1560
      ScaleHeight     =   43
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   43
      TabIndex        =   26
      Top             =   10440
      Width           =   645
   End
   Begin VB.PictureBox PicLettre 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "CG Omega"
         Size            =   20.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   645
      Index           =   3
      Left            =   2280
      ScaleHeight     =   43
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   43
      TabIndex        =   25
      Top             =   10440
      Width           =   645
   End
   Begin VB.PictureBox PicLettre 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "CG Omega"
         Size            =   20.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   645
      Index           =   4
      Left            =   3000
      ScaleHeight     =   43
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   43
      TabIndex        =   24
      Top             =   10440
      Width           =   645
   End
   Begin VB.PictureBox PicLettre 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "CG Omega"
         Size            =   20.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   645
      Index           =   5
      Left            =   3720
      ScaleHeight     =   43
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   43
      TabIndex        =   23
      Top             =   10440
      Width           =   645
   End
   Begin VB.PictureBox PicLettre 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "CG Omega"
         Size            =   20.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   645
      Index           =   6
      Left            =   4440
      ScaleHeight     =   43
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   43
      TabIndex        =   22
      Top             =   10440
      Width           =   645
   End
   Begin VB.PictureBox PicLettre 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "CG Omega"
         Size            =   20.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   645
      Index           =   7
      Left            =   5160
      ScaleHeight     =   43
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   43
      TabIndex        =   21
      Top             =   10440
      Width           =   645
   End
   Begin VB.PictureBox PicLettre 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "CG Omega"
         Size            =   20.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   645
      Index           =   8
      Left            =   5880
      ScaleHeight     =   43
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   43
      TabIndex        =   20
      Top             =   10440
      Width           =   645
   End
   Begin VB.PictureBox PicLettre 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "CG Omega"
         Size            =   20.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   645
      Index           =   9
      Left            =   6600
      ScaleHeight     =   43
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   43
      TabIndex        =   19
      Top             =   10440
      Width           =   645
   End
   Begin VB.PictureBox PicLettre 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "CG Omega"
         Size            =   20.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   645
      Index           =   10
      Left            =   7320
      ScaleHeight     =   43
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   43
      TabIndex        =   18
      Top             =   10440
      Width           =   645
   End
   Begin VB.PictureBox PicLettre 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "CG Omega"
         Size            =   20.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   645
      Index           =   11
      Left            =   8040
      ScaleHeight     =   43
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   43
      TabIndex        =   17
      Top             =   10440
      Width           =   645
   End
   Begin VB.PictureBox PicLettre 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "CG Omega"
         Size            =   20.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   645
      Index           =   12
      Left            =   8760
      ScaleHeight     =   43
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   43
      TabIndex        =   16
      Top             =   10440
      Width           =   645
   End
   Begin VB.PictureBox PicLettre 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "CG Omega"
         Size            =   20.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   645
      Index           =   13
      Left            =   120
      ScaleHeight     =   43
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   43
      TabIndex        =   15
      Top             =   11160
      Width           =   645
   End
   Begin VB.PictureBox PicLettre 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "CG Omega"
         Size            =   20.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   645
      Index           =   14
      Left            =   840
      ScaleHeight     =   43
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   43
      TabIndex        =   14
      Top             =   11160
      Width           =   645
   End
   Begin VB.PictureBox PicLettre 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "CG Omega"
         Size            =   20.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   645
      Index           =   15
      Left            =   1560
      ScaleHeight     =   43
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   43
      TabIndex        =   13
      Top             =   11160
      Width           =   645
   End
   Begin VB.PictureBox PicLettre 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "CG Omega"
         Size            =   20.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   645
      Index           =   16
      Left            =   2280
      ScaleHeight     =   43
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   43
      TabIndex        =   12
      Top             =   11160
      Width           =   645
   End
   Begin VB.PictureBox PicLettre 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "CG Omega"
         Size            =   20.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   645
      Index           =   17
      Left            =   3000
      ScaleHeight     =   43
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   43
      TabIndex        =   11
      Top             =   11160
      Width           =   645
   End
   Begin VB.PictureBox PicLettre 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "CG Omega"
         Size            =   20.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   645
      Index           =   18
      Left            =   3720
      ScaleHeight     =   43
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   43
      TabIndex        =   10
      Top             =   11160
      Width           =   645
   End
   Begin VB.PictureBox PicLettre 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "CG Omega"
         Size            =   20.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   645
      Index           =   19
      Left            =   4440
      ScaleHeight     =   43
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   43
      TabIndex        =   9
      Top             =   11160
      Width           =   645
   End
   Begin VB.PictureBox PicLettre 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "CG Omega"
         Size            =   20.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   645
      Index           =   20
      Left            =   5160
      ScaleHeight     =   43
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   43
      TabIndex        =   8
      Top             =   11160
      Width           =   645
   End
   Begin VB.PictureBox PicLettre 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "CG Omega"
         Size            =   20.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   645
      Index           =   21
      Left            =   5880
      ScaleHeight     =   43
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   43
      TabIndex        =   7
      Top             =   11160
      Width           =   645
   End
   Begin VB.PictureBox PicLettre 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "CG Omega"
         Size            =   20.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   645
      Index           =   22
      Left            =   6600
      ScaleHeight     =   43
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   43
      TabIndex        =   6
      Top             =   11160
      Width           =   645
   End
   Begin VB.PictureBox PicLettre 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "CG Omega"
         Size            =   20.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   645
      Index           =   23
      Left            =   7320
      ScaleHeight     =   43
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   43
      TabIndex        =   5
      Top             =   11160
      Width           =   645
   End
   Begin VB.PictureBox PicLettre 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "CG Omega"
         Size            =   20.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   645
      Index           =   24
      Left            =   8040
      ScaleHeight     =   43
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   43
      TabIndex        =   4
      Top             =   11160
      Width           =   645
   End
   Begin VB.PictureBox PicLettre 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "CG Omega"
         Size            =   20.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   645
      Index           =   25
      Left            =   8760
      ScaleHeight     =   43
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   43
      TabIndex        =   3
      Top             =   11160
      Width           =   645
   End
   Begin VB.PictureBox PicLettre 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "CG Omega"
         Size            =   20.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   645
      Index           =   0
      Left            =   120
      ScaleHeight     =   43
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   43
      TabIndex        =   2
      Top             =   10440
      Width           =   645
   End
   Begin VB.PictureBox PicLettre 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "CG Omega"
         Size            =   20.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   645
      Index           =   26
      Left            =   9480
      ScaleHeight     =   43
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   43
      TabIndex        =   1
      Top             =   11160
      Width           =   645
   End
   Begin VB.PictureBox Grille 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   9930
      Left            =   120
      ScaleHeight     =   662
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   662
      TabIndex        =   0
      Top             =   120
      Width           =   9930
      Begin VB.PictureBox Grille35 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   8130
         Left            =   0
         Picture         =   "FrmPics.frx":1A7C2C
         ScaleHeight     =   542
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   542
         TabIndex        =   33
         Top             =   1560
         Width           =   8130
      End
   End
End
Attribute VB_Name = "FrmPics"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False



