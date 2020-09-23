VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form frmChess3Offline 
   BackColor       =   &H00FF8080&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Ni-Star Enterprises - ChessMASTER Mark III: Offline Mode  - Niall Mooney"
   ClientHeight    =   6480
   ClientLeft      =   2145
   ClientTop       =   2715
   ClientWidth     =   9840
   Icon            =   "Chess3.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   432
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   656
   StartUpPosition =   1  'CenterOwner
   Begin VB.Timer Timer1 
      Interval        =   100
      Left            =   3525
      Top             =   1755
   End
   Begin VB.ComboBox cmbBacky 
      Height          =   315
      Left            =   120
      Style           =   2  'Dropdown List
      TabIndex        =   49
      Top             =   6000
      Width           =   3375
   End
   Begin VB.PictureBox picPiece 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      Height          =   780
      Index           =   1
      Left            =   7245
      Picture         =   "Chess3.frx":08CA
      ScaleHeight     =   48
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   48
      TabIndex        =   44
      Top             =   -600
      Visible         =   0   'False
      Width           =   780
   End
   Begin VB.PictureBox picPiece 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      Height          =   780
      Index           =   2
      Left            =   7605
      Picture         =   "Chess3.frx":240C
      ScaleHeight     =   48
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   48
      TabIndex        =   43
      Top             =   -600
      Visible         =   0   'False
      Width           =   780
   End
   Begin VB.PictureBox picPiece 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      Height          =   780
      Index           =   3
      Left            =   7965
      Picture         =   "Chess3.frx":3F4E
      ScaleHeight     =   48
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   48
      TabIndex        =   42
      Top             =   -600
      Visible         =   0   'False
      Width           =   780
   End
   Begin VB.PictureBox picPiece 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      Height          =   780
      Index           =   4
      Left            =   8325
      Picture         =   "Chess3.frx":5A90
      ScaleHeight     =   48
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   48
      TabIndex        =   41
      Top             =   -600
      Visible         =   0   'False
      Width           =   780
   End
   Begin VB.PictureBox picPiece 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      Height          =   780
      Index           =   5
      Left            =   8685
      Picture         =   "Chess3.frx":75D2
      ScaleHeight     =   48
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   48
      TabIndex        =   40
      Top             =   -600
      Visible         =   0   'False
      Width           =   780
   End
   Begin VB.PictureBox picPiece 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      Height          =   780
      Index           =   6
      Left            =   9045
      Picture         =   "Chess3.frx":9114
      ScaleHeight     =   48
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   48
      TabIndex        =   39
      Top             =   -600
      Visible         =   0   'False
      Width           =   780
   End
   Begin VB.PictureBox picPiece 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      Height          =   780
      Index           =   11
      Left            =   7245
      Picture         =   "Chess3.frx":AC56
      ScaleHeight     =   48
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   48
      TabIndex        =   38
      Top             =   -360
      Visible         =   0   'False
      Width           =   780
   End
   Begin VB.PictureBox picPiece 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      Height          =   780
      Index           =   12
      Left            =   7605
      Picture         =   "Chess3.frx":C798
      ScaleHeight     =   48
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   48
      TabIndex        =   37
      Top             =   -360
      Visible         =   0   'False
      Width           =   780
   End
   Begin VB.PictureBox picPiece 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      Height          =   780
      Index           =   13
      Left            =   7965
      Picture         =   "Chess3.frx":E2DA
      ScaleHeight     =   48
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   48
      TabIndex        =   36
      Top             =   -360
      Visible         =   0   'False
      Width           =   780
   End
   Begin VB.PictureBox picPiece 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      Height          =   780
      Index           =   14
      Left            =   8325
      Picture         =   "Chess3.frx":FE1C
      ScaleHeight     =   48
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   48
      TabIndex        =   35
      Top             =   -360
      Visible         =   0   'False
      Width           =   780
   End
   Begin VB.PictureBox picPiece 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      Height          =   780
      Index           =   15
      Left            =   8685
      Picture         =   "Chess3.frx":1195E
      ScaleHeight     =   48
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   48
      TabIndex        =   34
      Top             =   -360
      Visible         =   0   'False
      Width           =   780
   End
   Begin VB.PictureBox picPiece 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      Height          =   780
      Index           =   16
      Left            =   9045
      Picture         =   "Chess3.frx":134A0
      ScaleHeight     =   48
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   48
      TabIndex        =   33
      Top             =   -360
      Visible         =   0   'False
      Width           =   780
   End
   Begin VB.PictureBox picPiece 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      Height          =   780
      Index           =   21
      Left            =   7245
      Picture         =   "Chess3.frx":14FE2
      ScaleHeight     =   48
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   48
      TabIndex        =   32
      Top             =   -120
      Visible         =   0   'False
      Width           =   780
   End
   Begin VB.PictureBox picPiece 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      Height          =   780
      Index           =   22
      Left            =   7605
      Picture         =   "Chess3.frx":151AC
      ScaleHeight     =   48
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   48
      TabIndex        =   31
      Top             =   -120
      Visible         =   0   'False
      Width           =   780
   End
   Begin VB.PictureBox picPiece 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      Height          =   780
      Index           =   23
      Left            =   7965
      Picture         =   "Chess3.frx":15376
      ScaleHeight     =   48
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   48
      TabIndex        =   30
      Top             =   -120
      Visible         =   0   'False
      Width           =   780
   End
   Begin VB.PictureBox picPiece 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      Height          =   780
      Index           =   24
      Left            =   8325
      Picture         =   "Chess3.frx":15540
      ScaleHeight     =   48
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   48
      TabIndex        =   29
      Top             =   -120
      Visible         =   0   'False
      Width           =   780
   End
   Begin VB.PictureBox picPiece 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      Height          =   780
      Index           =   25
      Left            =   8685
      Picture         =   "Chess3.frx":1570A
      ScaleHeight     =   48
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   48
      TabIndex        =   28
      Top             =   -120
      Visible         =   0   'False
      Width           =   780
   End
   Begin VB.PictureBox picPiece 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      Height          =   780
      Index           =   26
      Left            =   9045
      Picture         =   "Chess3.frx":158D4
      ScaleHeight     =   48
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   48
      TabIndex        =   27
      Top             =   -120
      Visible         =   0   'False
      Width           =   780
   End
   Begin MSComDlg.CommonDialog CD 
      Left            =   840
      Top             =   480
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
      DialogTitle     =   "ChessMASTER NiFile Algorithm"
      Filter          =   "Chess Board NiFile|*.Nif"
   End
   Begin VB.CommandButton Command6 
      Caption         =   "Resign / Throw in towel"
      Height          =   375
      Left            =   120
      TabIndex        =   26
      Top             =   4080
      Width           =   3375
   End
   Begin VB.CommandButton Command5 
      Caption         =   "Redraw Board"
      Height          =   375
      Left            =   120
      TabIndex        =   25
      Top             =   4560
      Width           =   3375
   End
   Begin VB.CommandButton Command4 
      Caption         =   "Save Game"
      Height          =   375
      Left            =   120
      TabIndex        =   24
      Top             =   3600
      Width           =   3375
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Load Game to Edit"
      Height          =   375
      Left            =   120
      TabIndex        =   23
      Top             =   3120
      Width           =   3375
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Load Game and Start Game"
      Height          =   375
      Left            =   120
      TabIndex        =   22
      Top             =   2520
      Width           =   3375
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Start Chess Match"
      Height          =   375
      Left            =   120
      TabIndex        =   18
      Top             =   2040
      Width           =   3375
   End
   Begin VB.ComboBox cmbPieces 
      Height          =   315
      Left            =   120
      Locked          =   -1  'True
      Style           =   2  'Dropdown List
      TabIndex        =   17
      Top             =   5280
      Width           =   3375
   End
   Begin VB.ComboBox cmbBoard 
      Height          =   315
      Left            =   120
      Style           =   2  'Dropdown List
      TabIndex        =   16
      Top             =   5640
      Width           =   3375
   End
   Begin VB.PictureBox picEffector 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00FF0000&
      Height          =   780
      Index           =   1
      Left            =   6360
      Picture         =   "Chess3.frx":15A9E
      ScaleHeight     =   48
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   48
      TabIndex        =   15
      Top             =   120
      Visible         =   0   'False
      Width           =   780
   End
   Begin VB.PictureBox picEffector 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00FF0000&
      Height          =   780
      Index           =   0
      Left            =   6360
      Picture         =   "Chess3.frx":167E0
      ScaleHeight     =   48
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   48
      TabIndex        =   14
      Top             =   -120
      Visible         =   0   'False
      Width           =   780
   End
   Begin VB.PictureBox picFrontBuffer 
      AutoRedraw      =   -1  'True
      Height          =   5820
      Left            =   3960
      ScaleHeight     =   384
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   384
      TabIndex        =   0
      Top             =   600
      Width           =   5820
   End
   Begin VB.Image imgBacky1 
      BorderStyle     =   1  'Fixed Single
      Enabled         =   0   'False
      Height          =   6540
      Left            =   3600
      Picture         =   "Chess3.frx":1745B
      Top             =   4200
      Visible         =   0   'False
      Width           =   9900
   End
   Begin VB.Image imgBacky0 
      BorderStyle     =   1  'Fixed Single
      Enabled         =   0   'False
      Height          =   6540
      Left            =   3720
      Picture         =   "Chess3.frx":20C65
      Top             =   4080
      Visible         =   0   'False
      Width           =   9900
   End
   Begin VB.Label lblFile 
      BackStyle       =   0  'Transparent
      Caption         =   "E"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   4
      Left            =   7200
      TabIndex        =   48
      Top             =   240
      Width           =   255
   End
   Begin VB.Label lblFile 
      BackStyle       =   0  'Transparent
      Caption         =   "F"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   5
      Left            =   7920
      TabIndex        =   47
      Top             =   240
      Width           =   255
   End
   Begin VB.Label lblFile 
      BackStyle       =   0  'Transparent
      Caption         =   "G"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   6
      Left            =   8640
      TabIndex        =   46
      Top             =   240
      Width           =   255
   End
   Begin VB.Label lblFile 
      BackStyle       =   0  'Transparent
      Caption         =   "H"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   7
      Left            =   9360
      TabIndex        =   45
      Top             =   240
      Width           =   255
   End
   Begin VB.Label Label4 
      BackStyle       =   0  'Transparent
      Caption         =   "ChessMASTER Mark II Graphic Sets"
      ForeColor       =   &H000000C0&
      Height          =   255
      Left            =   120
      TabIndex        =   21
      Top             =   5040
      Width           =   3255
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "ChessMASTER Artifical Intelligence"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000C0&
      Height          =   255
      Left            =   120
      TabIndex        =   20
      Top             =   1680
      Width           =   3375
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "ChessMASTER Mark III Game Mode"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000C0&
      Height          =   255
      Left            =   120
      TabIndex        =   19
      Top             =   1440
      Width           =   3375
   End
   Begin VB.Label lblFile 
      BackStyle       =   0  'Transparent
      Caption         =   "D"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   3
      Left            =   6435
      TabIndex        =   13
      Top             =   240
      Width           =   255
   End
   Begin VB.Label lblFile 
      BackStyle       =   0  'Transparent
      Caption         =   "C"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   2
      Left            =   5715
      TabIndex        =   12
      Top             =   240
      Width           =   255
   End
   Begin VB.Label lblFile 
      BackStyle       =   0  'Transparent
      Caption         =   "B"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   1
      Left            =   4995
      TabIndex        =   11
      Top             =   240
      Width           =   255
   End
   Begin VB.Label lblFile 
      BackStyle       =   0  'Transparent
      Caption         =   "A"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   0
      Left            =   4275
      TabIndex        =   10
      Top             =   240
      Width           =   255
   End
   Begin VB.Label lblRank 
      BackStyle       =   0  'Transparent
      Caption         =   "8"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   7
      Left            =   3720
      TabIndex        =   9
      Top             =   5880
      Width           =   255
   End
   Begin VB.Label lblRank 
      BackStyle       =   0  'Transparent
      Caption         =   "7"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   6
      Left            =   3720
      TabIndex        =   8
      Top             =   5160
      Width           =   255
   End
   Begin VB.Label lblRank 
      BackStyle       =   0  'Transparent
      Caption         =   "6"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   5
      Left            =   3720
      TabIndex        =   7
      Top             =   4440
      Width           =   255
   End
   Begin VB.Label lblRank 
      BackStyle       =   0  'Transparent
      Caption         =   "5"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   4
      Left            =   3720
      TabIndex        =   6
      Top             =   3720
      Width           =   255
   End
   Begin VB.Label lblRank 
      BackStyle       =   0  'Transparent
      Caption         =   "4"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   3
      Left            =   3720
      TabIndex        =   5
      Top             =   3000
      Width           =   255
   End
   Begin VB.Label lblRank 
      BackStyle       =   0  'Transparent
      Caption         =   "3"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   2
      Left            =   3720
      TabIndex        =   4
      Top             =   2280
      Width           =   255
   End
   Begin VB.Label lblRank 
      BackStyle       =   0  'Transparent
      Caption         =   "2"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   1
      Left            =   3720
      TabIndex        =   3
      Top             =   1560
      Width           =   255
   End
   Begin VB.Label lblRank 
      BackStyle       =   0  'Transparent
      Caption         =   "1"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   0
      Left            =   3720
      TabIndex        =   2
      Top             =   840
      Width           =   255
   End
   Begin VB.Label lblStatus 
      Alignment       =   2  'Center
      BackColor       =   &H00FF8080&
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Caption         =   "No Game Currently in Session"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   1215
      Left            =   120
      TabIndex        =   1
      Top             =   120
      Width           =   3375
   End
   Begin VB.Menu mnuFile 
      Caption         =   "&File"
      Begin VB.Menu mnuExit 
         Caption         =   "Exit"
      End
   End
   Begin VB.Menu mnuAIType 
      Caption         =   "AI Type"
      Visible         =   0   'False
      Begin VB.Menu mnuAI 
         Caption         =   "PChessMASTER WoodPusher"
         Index           =   0
      End
      Begin VB.Menu mnuAI 
         Caption         =   "ChessMASTER Mark I (2000, NSE4)"
         Index           =   1
      End
      Begin VB.Menu mnuAI 
         Caption         =   "ChessMASTER Mark II (2001, NSE5)"
         Checked         =   -1  'True
         Index           =   2
      End
      Begin VB.Menu mnuAI 
         Caption         =   "ChessMASTER Mark III (2003, NSE7)"
         Index           =   3
      End
   End
   Begin VB.Menu mnuGametype 
      Caption         =   "Game Type"
      Visible         =   0   'False
      Begin VB.Menu mnuGamePlayers 
         Caption         =   "Player Versus Player"
         Checked         =   -1  'True
         Index           =   0
      End
      Begin VB.Menu mnuGamePlayers 
         Caption         =   "Player Versus CPU-AI"
         Index           =   1
      End
      Begin VB.Menu mnuGamePlayers 
         Caption         =   "CPU-AI Versus CPU-AI"
         Index           =   2
      End
   End
   Begin VB.Menu mnuGraphics 
      Caption         =   "Graphics"
      Begin VB.Menu mnuDangerShow 
         Caption         =   "Show Units in Danger"
         Checked         =   -1  'True
      End
   End
   Begin VB.Menu mnuOptions 
      Caption         =   "Options"
      Begin VB.Menu mnuAImnu 
         Caption         =   "Artifical Intelligence"
         Begin VB.Menu mnuDebugVerbose 
            Caption         =   "Debugging Verbose"
         End
         Begin VB.Menu mnuAIMultiTask 
            Caption         =   "Allow MultiTasking"
            Checked         =   -1  'True
         End
      End
      Begin VB.Menu mnuViewLog 
         Caption         =   "View Match Log"
      End
   End
   Begin VB.Menu mnuAbout 
      Caption         =   "About"
   End
End
Attribute VB_Name = "frmChess3Offline"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Ni-Star Enterprises
'ChessMASTER Mark III Core Class Integrated
'NSE7h - Provisional ChessMASTER Mark III
'Niall Mooney 23/2/2002

Option Explicit

Const bRook = 1
Const bKnight = 2
Const bBishop = 3
Const bQueen = 4
Const bKing = 5
Const bPawn = 6

Const wRook = 11
Const wKnight = 12
Const wBishop = 13
Const wQueen = 14
Const wKing = 15
Const wPawn = 16

Const LineDraw = 0
Const UFBox = 1
Const FBox = 2
Const UFEllipse = 3
Const FEllipse = 4
Const Star = 5

Dim Lx As Integer, Ly As Integer, SelX As Integer, SelY As Integer, Selected As Boolean
Dim Board(7, 7) As Byte
Dim BoardAI() As Byte
Dim Turn As Integer
Dim AIType As Integer
Dim AIToMove As Boolean
Dim MenuChoiceRate As Single
Dim PriorityMenu As Integer
Dim Mode As String 'OfflinePP, OfflinePC, OfflineCC, Client, Server
Dim DrawTool As Integer
Dim MakeMoveRecord As String 'Debugging

Private Declare Function BitBlt Lib "gdi32" (ByVal hDestDC As Long, ByVal X As Long, ByVal Y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal XSrc As Long, ByVal YSrc As Long, ByVal dwRop As Long) As Long
Private Declare Function GetTickCount Lib "kernel32" () As Long
Private Declare Function AlphaBlending Lib "msimg32.dll" Alias "AlphaBlend" (ByVal hdcDest As Long, ByVal nXOriginDest As Long, ByVal nYOriginDest As Long, ByVal nWidthDest As Long, ByVal nHeightDest As Long, ByVal hdcSrc As Long, ByVal nXOriginSrc As Long, ByVal nYOriginSrc As Long, ByVal nWidthSrc As Long, ByVal nHeightSrc As Long, ByVal BF As Long) As Long
Private Declare Function DrawTransparent Lib "msimg32.dll" Alias "TransparentBlt" (ByVal hdcDest As Long, ByVal nXOriginDest As Long, ByVal nYOriginDest As Long, ByVal nWidthDest As Long, ByVal nHeightDest As Long, ByVal hdcSrc As Long, ByVal nXOriginSrc As Long, ByVal nYOriginSrc As Long, ByVal nWidthSrc As Long, ByVal nHeightSrc As Long, ByVal crTransparent As Long) As Long

Private Sub cmbBacky_Click()
Select Case cmbBacky.ListIndex
Case 0: Me.Picture = imgBacky0.Picture
Case 1: Me.Picture = imgBacky1.Picture
End Select
End Sub

Private Sub cmbBoard_Click()
Select Case cmbBoard.ListIndex
Case 0: picFrontBuffer.Picture = LoadResPicture(101, 0)
Case 1: picFrontBuffer.Picture = LoadResPicture(102, 0)
Case 2: picFrontBuffer.Picture = LoadResPicture(103, 0)
Case 3: picFrontBuffer.Picture = LoadResPicture(104, 0)
Case 4: picFrontBuffer.Picture = LoadResPicture(105, 0)
Case 5: picFrontBuffer.Picture = LoadResPicture(106, 0)
Case 6: picFrontBuffer.Picture = LoadResPicture(107, 0)
Case 7: picFrontBuffer.Picture = LoadResPicture(108, 0)
Case 8: picFrontBuffer.Picture = LoadResPicture(109, 0)
Case 9: picFrontBuffer.Picture = LoadResPicture(110, 0)
Case 10: picFrontBuffer.Picture = LoadResPicture(111, 0)
Case 11: picFrontBuffer.Picture = LoadResPicture(112, 0)
Case 12: picFrontBuffer.Picture = LoadResPicture(113, 0)
End Select
Call RenderBoard
End Sub

Private Sub cmbPieces_Change()
Call RenderBoard
End Sub

Private Sub Command1_Click()
If mnuAI(0).Checked Then AIType = 0
If mnuAI(1).Checked Then AIType = 1
If mnuAI(2).Checked Then AIType = 2
If mnuAI(3).Checked Then AIType = 3

If mnuGamePlayers(0).Checked Then Mode = "OfflinePP"
If mnuGamePlayers(1).Checked Then Mode = "OfflinePC"
If mnuGamePlayers(2).Checked Then Mode = "OfflineCC"

InitiateNewGame
End Sub

Private Sub Command2_Click()
On Error Resume Next

If Mode <> "" Then Exit Sub

Dim J As String
Dim X&, Y&

CD.FileName = ""
CD.ShowOpen

If CD.FileName = "*.Nif" Or CD.FileName = "" Then Exit Sub
If UCase$(Right$(CD.FileName, 4)) <> ".NIF" Then Exit Sub

Open CD.FileName For Input As #5
Line Input #5, J$
Mode = J$
Line Input #5, J$
Turn = CByte(Val(J$))

For X = 0 To 7
For Y = 0 To 7
Line Input #5, J$
Board(X, Y) = CByte(Val(J$))
Next
Next
Close #5

If Mode = "OfflineCC" Then PerformAI

If Err Then MsgBox Err.Description, vbInformation, "ChessMASTER NiFile Algorithm"
End Sub

Private Sub Command3_Click()
If Mode <> "" Then Exit Sub

On Error Resume Next

Dim J As String
Dim X&, Y&

CD.FileName = ""
CD.ShowOpen

If CD.FileName = "*.Nif" Or CD.FileName = "" Then Exit Sub
If UCase$(Right$(CD.FileName, 4)) <> ".NIF" Then Exit Sub

Open CD.FileName For Input As #5
Line Input #5, J$
Mode = J$
Line Input #5, J$
Turn = CByte(Val(J$))

For X = 0 To 7
For Y = 0 To 7
Line Input #5, J$
Board(X, Y) = CByte(Val(J$))
Next
Next
Close #5

Mode = "EditMode"
RenderBoard
If Err Then MsgBox Err.Description, vbCritical, "ChessMASTER NiFile Algorithm"
End Sub

Private Sub Command4_Click()
'On Error Resume Next

Dim X&, Y&

CD.FileName = ""
CD.ShowSave
If CD.FileName = "" Then Exit Sub

Open CD.FileName For Output As #5
Print #5, Mode
Print #5, Trim$(Str$(Turn))
For X = 0 To 7
For Y = 0 To 7
Print #5, Trim$(Str$(Board(X, Y)))
Next
Next
Close #5

'If Err Then MsgBox Err.Description, vbCritical, "ChessMASTER"

Call RenderBoard
End Sub

Private Sub Command5_Click()
Call RenderBoard
End Sub

Private Sub Command6_Click()
Select Case Mode
Case "OfflinePP"
If Turn = 0 Then
If MsgBox("Black Player are you sure that you wish to throw in the towel and afford victory to the Enemy White?", vbQuestion + vbYesNo, "NSE ChessMASTER Mark III") = vbYes Then
EndGame
End If
End If
If Turn = 1 Then
If MsgBox("White Player are you sure that you wish to throw in the towel and afford victory to the Enemy Black?", vbQuestion + vbYesNo, "NSE ChessMASTER Mark III") = vbYes Then
EndGame
End If
End If
Case "OfflinePC"
If Turn = 1 Then
If MsgBox("Player are you sure that you wish to throw in the towel and afford victory to the Enemy CPU?", vbQuestion + vbYesNo, "NSE ChessMASTER Mark III") = vbYes Then
EndGame
End If
End If
Case "OfflineCC"
If MsgBox("Are you sure that you wish to end the game?", vbApplicationModal + vbQuestion + vbYesNo, "ChessMASTER Mark III") = vbYes Then
EndGame
End If
End Select
End Sub

Private Sub Form_Initialize()
MenuChoiceRate = 1.67
AIType = 1

cmbPieces.AddItem "CM2 (RJSoft)"
cmbPieces.ListIndex = 0

cmbBoard.AddItem "NSE ChessMASTER"
cmbBoard.AddItem "NSE ChessMASTER Mk. II"
cmbBoard.AddItem "Green Emerald (RJSoft)"
cmbBoard.AddItem "Blue over white (RJSoft)"
cmbBoard.AddItem "Black Marble (RJSoft)"
cmbBoard.AddItem "Green Marble (RJSoft)"
cmbBoard.AddItem "Purple Stone (RJSoft)"
cmbBoard.AddItem "Brown Stone (RJSoft)"
cmbBoard.AddItem "NSE Agree Board 1"
cmbBoard.AddItem "NSE Agree Board 2"
cmbBoard.AddItem "NSE Agree Board 3"
cmbBoard.AddItem "NSE Agree Board 4"
cmbBoard.AddItem "NSE Agree Board 5"
cmbBoard.ListIndex = 1

cmbBacky.AddItem "Blues Skies Go Greener"
cmbBacky.AddItem "Rainbow Go Bizzarre"
cmbBacky.ListIndex = 0

Me.Picture = imgBacky0.Picture
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
Dim Ans As Long

Ans = MsgBox("Are you certain that you wish to leave at this time", vbQuestion + vbYesNo, "NSE Provisional Chess 3")
If Ans = vbYes Then
Unload Me
End
ElseIf Ans = vbNo Then
Cancel = 1
End If
End Sub

Public Sub InitiateNewGame()
Dim X As Long
Dim Y As Long

For X = 0 To 7
For Y = 0 To 7
Board(X, Y) = 0
Next Y
Next X

Board(0, 0) = 1
Board(1, 0) = 2
Board(2, 0) = 3
Board(3, 0) = 4
Board(4, 0) = 5
Board(5, 0) = 3
Board(6, 0) = 2
Board(7, 0) = 1

Board(0, 1) = 6
Board(1, 1) = 6
Board(2, 1) = 6
Board(3, 1) = 6
Board(4, 1) = 6
Board(5, 1) = 6
Board(6, 1) = 6
Board(7, 1) = 6

Board(0, 7) = 11
Board(1, 7) = 12
Board(2, 7) = 13
Board(3, 7) = 14
Board(4, 7) = 15
Board(5, 7) = 13
Board(6, 7) = 12
Board(7, 7) = 11

Board(0, 6) = 16
Board(1, 6) = 16
Board(2, 6) = 16
Board(3, 6) = 16
Board(4, 6) = 16
Board(5, 6) = 16
Board(6, 6) = 16
Board(7, 6) = 16

Turn = 1

RenderBoard

If Mode = "OfflineCC" Then PerformAI
End Sub

Public Sub MakeMove(Rank1 As Integer, File1 As Integer, Rank2 As Integer, File2 As Integer)

Call LogMove(MakeMoveRecord, Rank1, File1, Rank2, File2)

If Mode = "EditGame" Then
Board(Rank2, File2) = Board(Rank1, File1)
Board(Rank1, File1) = 0
If Turn = 0 Then Turn = 1 Else Turn = 0
Selected = False
RenderBoard
Exit Sub
End If

If MoveAlright(Rank1, File1, Rank2, File2) Then
Board(Rank2, File2) = Board(Rank1, File1)
Board(Rank1, File1) = 0
If Turn = 0 Then Turn = 1 Else Turn = 0
Selected = False
RenderBoard
End If

Call CheckGame
End Sub

Private Function MoveAlright(Sx As Integer, Sy As Integer, tx As Integer, ty As Integer) As Boolean
'On Error Resume Next
Dim Legal As Integer
Dim Colour As String
Dim ColourT As String
Dim Adder As Integer
Dim Taking As Boolean
Dim Moving As Boolean
Dim X&, Y&

If Board(Sx, Sy) <> 0 And Board(Sx, Sy) < 10 Then Colour$ = "B"
If Board(Sx, Sy) <> 0 And Board(Sx, Sy) > 10 Then Colour$ = "W"
If Board(Sx, Sy) <> 0 And Board(tx, ty) < 10 And Board(tx, ty) > 0 Then ColourT = "B"
If Board(Sx, Sy) <> 0 And Board(tx, ty) > 10 Then ColourT = "W"

If Colour = "W" Then Adder = 10

If Board(Sx, Sy) = 0 Or tx > 7 Or ty > 7 Or (Sx = tx And Sy = ty) Or Colour$ = ColourT Then    ' Invalid move
    MoveAlright = False
    Exit Function
Else
    Legal = 0
   
    If Board(Sx, Sy) = bKing + Adder And CheckPosition(tx, ty, Board(Sx, Sy)) Then
        MoveAlright = False
        Exit Function
    End If
    If Board(Sx, Sy) = bKing + Adder And CheckPosition(tx, ty, Board(Sx, Sy)) Then
        MoveAlright = False
        Exit Function
    End If
    
    'See if were trying to take a piece
    If Colour$ <> ColourT And ColourT <> "" Then
        Taking = True
        Moving = False
    Else ' We are just moving to a blank space.
        Taking = False
        Moving = True
    End If
    
    Select Case Board(Sx, Sy)
    Case bPawn + Adder 'Prawn movement
        If Taking Then
            If ty = Sy - 1 And Colour$ = "W" And (tx = Sx - 1 Or tx = Sx + 1) Then Legal = 5  'Black Up + Left/Right to White Taking a piece
            If ty = Sy + 1 And Colour$ = "B" And (tx = Sx - 1 Or tx = Sx + 1) Then Legal = 5  'White Up + Left/Right to Black Taking a piece
        Else
            If (Board(tx, ty) = 0 And ty = Sy - 1 And Colour$ = "W" And Sx = tx) Then Legal = 1    'White vertical
            If (Board(tx, ty) = 0 And ty = Sy + 1 And Colour$ = "B" And Sx = tx) Then Legal = 3    'Black Vertical
            If (Colour$ = "W" And Sy = 6 And ty = Sy - 2) And tx = Sx Then Legal = 1  'First move may be double WHITE
            If (Colour$ = "B" And Sy = 1 And ty = Sy + 2) And tx = Sx Then Legal = 3  'First move may be double BLACK
        End If

    Case bKing + Adder   'King movement
        If tx <= Sx + 1 And tx >= Sx - 1 And ty <= Sy + 1 And ty >= Sy - 1 Then Legal = 5 'Move in any direction by one
        
        ' Allow castling to the right
        If tx < 7 Then _
            If Board(tx + 1, ty) = bRook + Adder And _
                tx = Sx + 2 And ty = Sy Then Legal = 7
            
        ' Allow castling to the Left
        If tx > 2 Then _
            If Board(tx - 2, ty) = bRook + Adder And _
                tx = Sx - 2 And ty = Sy Then Legal = 7
            
    Case bQueen + Adder
        'Diagonal
        If Abs(Sx - tx) = Abs(Sy - ty) Or Abs(tx - Sx) = Abs(ty - Sy) Then Legal = 5
        'Left
        If tx < Sx And ty = Sy Then Legal = 4
        'Right
        If tx > Sx And ty = Sy Then Legal = 2
        'Up
        If tx = Sx And ty < Sy Then Legal = 1
        'Down
        If tx = Sx And ty > Sy Then Legal = 3
        
    Case bBishop + Adder
        If Abs(Sx - tx) = Abs(Sy - ty) Or Abs(tx - Sx) = Abs(ty - Sy) Then Legal = 5
        
    Case bKnight + Adder 'Horsey (Knight)
        If (tx = Sx + 2 And ty = Sy + 1) Then Legal = 6
        If (tx = Sx + 1 And ty = Sy + 2) Then Legal = 6
        If (tx = Sx - 2 And ty = Sy - 1) Then Legal = 6
        If (tx = Sx - 1 And ty = Sy - 2) Then Legal = 6
        If (tx = Sx + 2 And ty = Sy - 1) Then Legal = 6
        If (tx = Sx + 1 And ty = Sy - 2) Then Legal = 6
        If (tx = Sx - 2 And ty = Sy + 1) Then Legal = 6
        If (tx = Sx - 1 And ty = Sy + 2) Then Legal = 6
    
    Case bRook + Adder 'Castle (Rook)
        'Left
        If tx < Sx And ty = Sy Then Legal = 4
        'Right
        If tx > Sx And ty = Sy Then Legal = 2
        'Up
        If tx = Sx And ty < Sy Then Legal = 1
        'Down
        If tx = Sx And ty > Sy Then Legal = 3
    
    End Select
     'Check to see if we can move, without jumping a piece, based on Legal values
    
    Select Case Legal
    Case 0
        MoveAlright = False
        Exit Function
    Case 1 'Up
        For Y = Sy - 1 To ty Step -1
            If Y < 0 Then
                MoveAlright = False
                Exit Function
            End If
            
            If Moving And Board(Sx, Y) <> 0 Then
                MoveAlright = False
                Exit Function
            End If

            ' are we taking, is the current square blank, if its not blank is it the desired one
            If Taking And (Board(Sx, Y) <> 0 And (Sx <> tx Or Y <> ty)) Then
                MoveAlright = False
                Exit Function
            End If
        Next
    Case 2 'Right
        For X = Sx + 1 To tx Step 1
            If X < 0 Then
                MoveAlright = False
                Exit Function
            End If

            If Moving And Board(X, Sy) <> 0 Then
                MoveAlright = False
                Exit Function
            End If

            If Taking And (Board(X, Sy) <> 0 And (X <> tx Or Sy <> ty)) Then
                MoveAlright = False
                Exit Function
            End If

        Next
    Case 3 'Down
        For Y = Sy + 1 To ty Step 1
            If Y < 0 Then
                MoveAlright = False
                Exit Function
            End If

            If Moving And Board(Sx, Y) <> 0 Then
                MoveAlright = False
                Exit Function
            End If

            If Taking And (Board(Sx, Y) <> 0 And (Sx <> tx Or Y <> ty)) Then
                MoveAlright = False
                Exit Function
            End If

        Next
    Case 4 'Left
        For X = Sx - 1 To tx Step -1
            If X < 0 Then
                MoveAlright = False
                Exit Function
            End If

            If Moving And Board(X, Sy) <> 0 Then
                MoveAlright = False
                Exit Function
            End If

            If Taking And (Board(X, Sy) <> 0 And (X <> tx Or Sy <> ty)) Then
                MoveAlright = False
                Exit Function
            End If

        Next
    Case 5 'Diagonal
    If Sx > tx And Sy > ty Then 'Up Left
        For X = Sx - 1 To tx Step -1
            For Y = Sy - 1 To ty Step -1
                If Sx - X = Sy - Y Then 'Only check if it's a diagonal
                    If Moving And Board(X, Y) <> 0 Then
                        MoveAlright = False
                        Exit Function
                    End If

                    If Taking And Board(X, Y) <> 0 And (X <> tx Or Y <> ty) Then
                        MoveAlright = False
                        Exit Function
                    End If

                End If
            Next
        Next
    End If
    
    If Sx < tx And Sy > ty Then 'Up Right
        For X = Sx + 1 To tx
            For Y = Sy - 1 To ty Step -1
                If Sx - X = Y - Sy Then  'Only check if it's a diagonal
                    If Moving And Board(X, Y) <> 0 Then
                        MoveAlright = False
                        Exit Function
                    End If

                    If Taking And (Board(X, Y) <> 0 And (X <> tx Or Y <> ty)) Then
                        MoveAlright = False
                        Exit Function
                    End If

                End If
            Next
        Next
    End If
    If Sx > tx And Sy < ty Then 'Down Left
        For X = Sx - 1 To tx Step -1
            For Y = Sy + 1 To ty
                If X - Sx = Sy - Y Then 'Only check if it's a diagonal
                    If Moving And Board(X, Y) <> 0 Then
                        MoveAlright = False
                        Exit Function
                    End If

                    If Taking And (Board(X, Y) <> 0 And (X <> tx Or Y <> ty)) Then
                        MoveAlright = False
                        Exit Function
                    End If

                End If
            Next
        Next
    End If
    If Sx < tx And Sy < ty Then 'Down Right
        For X = Sx + 1 To tx Step 1
            For Y = Sy + 1 To ty Step 1
                If Sx - X = Sy - Y Then 'Only check if it's a diagonal
                    If Moving And Board(X, Y) <> 0 Then
                        MoveAlright = False
                        Exit Function
                    End If
                    
                    If Taking And (Board(X, Y) <> 0 And (X <> tx Or Y <> ty)) Then
                        MoveAlright = False
                        Exit Function
                    End If

                End If
            Next
        Next
    End If
    Case 6
        'It's a horsey & they're allowed to jump pieces!
    Case 7
        ' Castling Rule
        If Taking Then
            ' Move illegal so just exit
            MoveAlright = False
            Exit Function
        ElseIf Moving Then
        ' Legal left
        For X = Sx - 1 To tx Step -1
            If Moving And Board(X, Sy) <> 0 Then
                MoveAlright = False
                Exit Function
            End If
        Next X
        
        For X = Sx + 1 To tx Step 1
            If Moving And Board(X, Sy) <> 0 Then
                MoveAlright = False
                Exit Function
            End If
        Next X
        
        End If
    End Select
    
    'Successful Move!
    MoveAlright = True
End If
End Function

Private Function MoveLegal(Sx As Integer, Sy As Integer, tx As Integer, ty As Integer) As Boolean
'On Error Resume Next
Dim Legal As Integer
Dim Colour As String
Dim ColourT As String
Dim Adder As Integer
Dim Taking As Boolean
Dim Moving As Boolean
Dim Y&, X&

If Board(Sx, Sy) < 10 Then Colour$ = "B"
If Board(Sx, Sy) > 10 Then Colour$ = "W"
If Board(tx, ty) < 10 And Board(tx, ty) > 0 Then ColourT = "B"
If Board(tx, ty) > 10 Then ColourT = "W"

If Colour = "W" Then Adder = 10

If Board(Sx, Sy) = 0 Or tx > 7 Or ty > 7 Or (Sx = tx And Sy = ty) Or Colour$ = ColourT Then    ' Invalid move
    MoveLegal = False
    Exit Function
Else
    Legal = 0

    'See if were trying to take a piece
    If Colour$ <> ColourT And ColourT <> "" Then
        Taking = True
        Moving = False
    Else ' We are just moving to a blank space.
        Taking = False
        Moving = True
    End If
    
    Select Case Board(Sx, Sy)
    Case bPawn + Adder 'Prawn movement
        If Taking Then
            If ty = Sy - 1 And Colour$ = "W" And (tx = Sx - 1 Or tx = Sx + 1) Then Legal = 5  'Black Up + Left/Right to White Taking a piece
            If ty = Sy + 1 And Colour$ = "B" And (tx = Sx - 1 Or tx = Sx + 1) Then Legal = 5  'White Up + Left/Right to Black Taking a piece
        Else
            If (Board(tx, ty) = 0 And ty = Sy - 1 And Colour$ = "W" And Sx = tx) Then Legal = 1    'White vertical
            If (Board(tx, ty) = 0 And ty = Sy + 1 And Colour$ = "B" And Sx = tx) Then Legal = 3    'Black Vertical
            If (Colour$ = "W" And Sy = 7 And ty = Sy - 2) And tx = Sx Then Legal = 1  'First move may be double WHITE
            If (Colour$ = "B" And Sy = 2 And ty = Sy + 2) And tx = Sx Then Legal = 3  'First move may be double BLACK
        End If

    Case bKing + Adder   'King movement
        If tx <= Sx + 1 And tx >= Sx - 1 And ty <= Sy + 1 And ty >= Sy - 1 Then Legal = 5 'Move in any direction by one
        
        ' Allow castling to the right
        If tx < 7 Then _
            If Board(tx + 1, ty) = bRook + Adder And _
                tx = Sx + 2 And ty = Sy Then Legal = 7
            
        ' Allow castling to the Left
        If tx > 2 Then _
            If Board(tx - 2, ty) = bRook + Adder And _
                tx = Sx - 2 And ty = Sy Then Legal = 7
            
    Case bQueen + Adder
        'Diagonal
        If Abs(Sx - tx) = Abs(Sy - ty) Or Abs(tx - Sx) = Abs(ty - Sy) Then Legal = 5
        'Left
        If tx < Sx And ty = Sy Then Legal = 4
        'Right
        If tx > Sx And ty = Sy Then Legal = 2
        'Up
        If tx = Sx And ty < Sy Then Legal = 1
        'Down
        If tx = Sx And ty > Sy Then Legal = 3
        
    Case bBishop + Adder
        If Abs(Sx - tx) = Abs(Sy - ty) Or Abs(tx - Sx) = Abs(ty - Sy) Then Legal = 5
        
    Case bKnight + Adder 'Horsey
        If (tx = Sx + 2 And ty = Sy + 1) Then Legal = 6
        If (tx = Sx + 1 And ty = Sy + 2) Then Legal = 6
        If (tx = Sx - 2 And ty = Sy - 1) Then Legal = 6
        If (tx = Sx - 1 And ty = Sy - 2) Then Legal = 6
        If (tx = Sx + 2 And ty = Sy - 1) Then Legal = 6
        If (tx = Sx + 1 And ty = Sy - 2) Then Legal = 6
        If (tx = Sx - 2 And ty = Sy + 1) Then Legal = 6
        If (tx = Sx - 1 And ty = Sy + 2) Then Legal = 6
    
    Case bRook + Adder 'Castle (Rook)
        'Left
        If tx < Sx And ty = Sy Then Legal = 4
        'Right
        If tx > Sx And ty = Sy Then Legal = 2
        'Up
        If tx = Sx And ty < Sy Then Legal = 1
        'Down
        If tx = Sx And ty > Sy Then Legal = 3
    
    End Select
     'Check to see if we can move, without jumping a piece, based on Legal values
    
    Select Case Legal
    Case 0
        MoveLegal = False
        Exit Function
    Case 1 'Up
        For Y = Sy - 1 To ty Step -1
            If Y < 0 Then
                MoveLegal = False
                Exit Function
            End If
            
            If Moving And Board(Sx, Y) <> 0 Then
                MoveLegal = False
                Exit Function
            End If

            ' are we taking, is the current square blank, if its not blank is it the desired one
            If Taking And (Board(Sx, Y) <> 0 And (Sx <> tx Or Y <> ty)) Then
                MoveLegal = False
                Exit Function
            End If
        Next
    Case 2 'Right
        For X = Sx + 1 To tx Step 1
            If X < 0 Then
                MoveLegal = False
                Exit Function
            End If

            If Moving And Board(X, Sy) <> 0 Then
                MoveLegal = False
                Exit Function
            End If

            If Taking And (Board(X, Sy) <> 0 And (X <> tx Or Sy <> ty)) Then
                MoveLegal = False
                Exit Function
            End If

        Next
    Case 3 'Down
        For Y = Sy + 1 To ty Step 1
            If Y < 0 Then
                MoveLegal = False
                Exit Function
            End If

            If Moving And Board(Sx, Y) <> 0 Then
                MoveLegal = False
                Exit Function
            End If

            If Taking And (Board(Sx, Y) <> 0 And (Sx <> tx Or Y <> ty)) Then
                MoveLegal = False
                Exit Function
            End If

        Next
    Case 4 'Left
        For X = Sx - 1 To tx Step -1
            If X < 0 Then
                MoveLegal = False
                Exit Function
            End If

            If Moving And Board(X, Sy) <> 0 Then
                MoveLegal = False
                Exit Function
            End If

            If Taking And (Board(X, Sy) <> 0 And (X <> tx Or Sy <> ty)) Then
                MoveLegal = False
                Exit Function
            End If

        Next
    Case 5 'Diagonal
    If Sx > tx And Sy > ty Then 'Up Left
        For X = Sx - 1 To tx Step -1
            For Y = Sy - 1 To ty Step -1
                If Sx - X = Sy - Y Then 'Only check if it's a diagonal
                    If Moving And Board(X, Y) <> 0 Then
                        MoveLegal = False
                        Exit Function
                    End If

                    If Taking And Board(X, Y) <> 0 And (X <> tx Or Y <> ty) Then
                        MoveLegal = False
                        Exit Function
                    End If

                End If
            Next
        Next
    End If
    
    If Sx < tx And Sy > ty Then 'Up Right
        For X = Sx + 1 To tx
            For Y = Sy - 1 To ty Step -1
                If Sx - X = Y - Sy Then  'Only check if it's a diagonal
                    If Moving And Board(X, Y) <> 0 Then
                        MoveLegal = False
                        Exit Function
                    End If

                    If Taking And (Board(X, Y) <> 0 And (X <> tx Or Y <> ty)) Then
                        MoveLegal = False
                        Exit Function
                    End If

                End If
            Next
        Next
    End If
    If Sx > tx And Sy < ty Then 'Down Left
        For X = Sx - 1 To tx Step -1
            For Y = Sy + 1 To ty
                If X - Sx = Sy - Y Then 'Only check if it's a diagonal
                    If Moving And Board(X, Y) <> 0 Then
                        MoveLegal = False
                        Exit Function
                    End If

                    If Taking And (Board(X, Y) <> 0 And (X <> tx Or Y <> ty)) Then
                        MoveLegal = False
                        Exit Function
                    End If

                End If
            Next
        Next
    End If
    If Sx < tx And Sy < ty Then 'Down Right
        For X = Sx + 1 To tx Step 1
            For Y = Sy + 1 To ty Step 1
                If Sx - X = Sy - Y Then 'Only check if it's a diagonal
                    If Moving And Board(X, Y) <> 0 Then
                        MoveLegal = False
                        Exit Function
                    End If
                    
                    If Taking And (Board(X, Y) <> 0 And (X <> tx Or Y <> ty)) Then
                        MoveLegal = False
                        Exit Function
                    End If

                End If
            Next
        Next
    End If
    Case 6
        'It's a horsey & they're allowed to jump pieces!
    Case 7
        ' Castling Rule
        If Taking Then
            ' Move illegal so just exit
            MoveLegal = False
            Exit Function
        ElseIf Moving Then
        ' Legal left
        For X = Sx - 1 To tx Step -1
            If Moving And Board(X, Sy) <> 0 Then
                MoveLegal = False
                Exit Function
            End If
        Next X
        
        For X = Sx + 1 To tx Step 1
            If Moving And Board(X, Sy) <> 0 Then
                MoveLegal = False
                Exit Function
            End If
        Next X
        
        End If
    End Select
    
    'Successful Move!
    MoveLegal = True
End If
End Function


Private Function MoveAlrightAI(Sx As Integer, Sy As Integer, tx As Integer, ty As Integer) As Boolean
'On Error Resume Next
Dim Legal As Integer
Dim Colour As String
Dim ColourT As String
Dim Adder As Integer
Dim Taking As Boolean
Dim Moving As Boolean
Dim X&, Y&

If BoardAI(Sx, Sy) <> 0 And BoardAI(Sx, Sy) < 10 Then Colour$ = "B"
If BoardAI(Sx, Sy) <> 0 And BoardAI(Sx, Sy) > 10 Then Colour$ = "W"
If BoardAI(Sx, Sy) <> 0 And BoardAI(tx, ty) < 10 And BoardAI(tx, ty) > 0 Then ColourT = "B"
If BoardAI(Sx, Sy) <> 0 And BoardAI(tx, ty) > 10 Then ColourT = "W"

If Colour = "W" Then Adder = 10

If BoardAI(Sx, Sy) = 0 Or tx > 7 Or ty > 7 Or (Sx = tx And Sy = ty) Or Colour$ = ColourT Then    ' Invalid move
    MoveAlrightAI = False
    Exit Function
Else
    Legal = 0

    If BoardAI(Sx, Sy) = bKing + Adder And CheckPositionAI(tx, ty) Then
        MoveAlrightAI = False
        Exit Function
    End If
    If BoardAI(Sx, Sy) <> bKing + Adder And CheckForCheck And CheckPosition(tx, ty) Then
        MoveAlrightAI = False
        Exit Function
    End If

    'See if were trying to take a piece
    If Colour$ <> ColourT And ColourT <> "" Then
        Taking = True
        Moving = False
    Else ' We are just moving to a blank space.
        Taking = False
        Moving = True
    End If
    
    Select Case BoardAI(Sx, Sy)
    Case bPawn + Adder 'Prawn movement
        If Taking Then
            If ty = Sy - 1 And Colour$ = "W" And (tx = Sx - 1 Or tx = Sx + 1) Then Legal = 5  'Black Up + Left/Right to White Taking a piece
            If ty = Sy + 1 And Colour$ = "B" And (tx = Sx - 1 Or tx = Sx + 1) Then Legal = 5  'White Up + Left/Right to Black Taking a piece
        Else
            If (BoardAI(tx, ty) = 0 And ty = Sy - 1 And Colour$ = "W" And Sx = tx) Then Legal = 1    'White vertical
            If (BoardAI(tx, ty) = 0 And ty = Sy + 1 And Colour$ = "B" And Sx = tx) Then Legal = 3    'Black Vertical
            If (Colour$ = "W" And Sy = 6 And ty = Sy - 2) And tx = Sx Then Legal = 1  'First move may be double WHITE
            If (Colour$ = "B" And Sy = 1 And ty = Sy + 2) And tx = Sx Then Legal = 3  'First move may be double BLACK
        End If

    Case bKing + Adder   'King movement
        If tx <= Sx + 1 And tx >= Sx - 1 And ty <= Sy + 1 And ty >= Sy - 1 Then Legal = 5 'Move in any direction by one
        
        ' Allow castling to the right
        If tx < 7 Then _
            If BoardAI(tx + 1, ty) = bRook + Adder And _
                tx = Sx + 2 And ty = Sy Then Legal = 7
            
        ' Allow castling to the Left
        If tx > 2 Then _
            If BoardAI(tx - 2, ty) = bRook + Adder And _
                tx = Sx - 2 And ty = Sy Then Legal = 7
            
    Case bQueen + Adder
        'Diagonal
        If Abs(Sx - tx) = Abs(Sy - ty) Or Abs(tx - Sx) = Abs(ty - Sy) Then Legal = 5
        'Left
        If tx < Sx And ty = Sy Then Legal = 4
        'Right
        If tx > Sx And ty = Sy Then Legal = 2
        'Up
        If tx = Sx And ty < Sy Then Legal = 1
        'Down
        If tx = Sx And ty > Sy Then Legal = 3
        
    Case bBishop + Adder
        If Abs(Sx - tx) = Abs(Sy - ty) Or Abs(tx - Sx) = Abs(ty - Sy) Then Legal = 5
        
    Case bKnight + Adder 'Horsey (Knight)
        If (tx = Sx + 2 And ty = Sy + 1) Then Legal = 6
        If (tx = Sx + 1 And ty = Sy + 2) Then Legal = 6
        If (tx = Sx - 2 And ty = Sy - 1) Then Legal = 6
        If (tx = Sx - 1 And ty = Sy - 2) Then Legal = 6
        If (tx = Sx + 2 And ty = Sy - 1) Then Legal = 6
        If (tx = Sx + 1 And ty = Sy - 2) Then Legal = 6
        If (tx = Sx - 2 And ty = Sy + 1) Then Legal = 6
        If (tx = Sx - 1 And ty = Sy + 2) Then Legal = 6
    
    Case bRook + Adder 'Castle (Rook)
        'Left
        If tx < Sx And ty = Sy Then Legal = 4
        'Right
        If tx > Sx And ty = Sy Then Legal = 2
        'Up
        If tx = Sx And ty < Sy Then Legal = 1
        'Down
        If tx = Sx And ty > Sy Then Legal = 3
    
    End Select
     'Check to see if we can move, without jumping a piece, based on Legal values
    
    Select Case Legal
    Case 0
        MoveAlrightAI = False
        Exit Function
    Case 1 'Up
        For Y = Sy - 1 To ty Step -1
            If Y < 0 Then
                MoveAlrightAI = False
                Exit Function
            End If
            
            If Moving And BoardAI(Sx, Y) <> 0 Then
                MoveAlrightAI = False
                Exit Function
            End If

            ' are we taking, is the current square blank, if its not blank is it the desired one
            If Taking And (BoardAI(Sx, Y) <> 0 And (Sx <> tx Or Y <> ty)) Then
                MoveAlrightAI = False
                Exit Function
            End If
        Next
    Case 2 'Right
        For X = Sx + 1 To tx Step 1
            If X < 0 Then
                MoveAlrightAI = False
                Exit Function
            End If

            If Moving And BoardAI(X, Sy) <> 0 Then
                MoveAlrightAI = False
                Exit Function
            End If

            If Taking And (BoardAI(X, Sy) <> 0 And (X <> tx Or Sy <> ty)) Then
                MoveAlrightAI = False
                Exit Function
            End If

        Next
    Case 3 'Down
        For Y = Sy + 1 To ty Step 1
            If Y < 0 Then
                MoveAlrightAI = False
                Exit Function
            End If

            If Moving And BoardAI(Sx, Y) <> 0 Then
                MoveAlrightAI = False
                Exit Function
            End If

            If Taking And (BoardAI(Sx, Y) <> 0 And (Sx <> tx Or Y <> ty)) Then
                MoveAlrightAI = False
                Exit Function
            End If

        Next
    Case 4 'Left
        For X = Sx - 1 To tx Step -1
            If X < 0 Then
                MoveAlrightAI = False
                Exit Function
            End If

            If Moving And BoardAI(X, Sy) <> 0 Then
                MoveAlrightAI = False
                Exit Function
            End If

            If Taking And (BoardAI(X, Sy) <> 0 And (X <> tx Or Sy <> ty)) Then
                MoveAlrightAI = False
                Exit Function
            End If

        Next
    Case 5 'Diagonal
    If Sx > tx And Sy > ty Then 'Up Left
        For X = Sx - 1 To tx Step -1
            For Y = Sy - 1 To ty Step -1
                If Sx - X = Sy - Y Then 'Only check if it's a diagonal
                    If Moving And BoardAI(X, Y) <> 0 Then
                        MoveAlrightAI = False
                        Exit Function
                    End If

                    If Taking And BoardAI(X, Y) <> 0 And (X <> tx Or Y <> ty) Then
                        MoveAlrightAI = False
                        Exit Function
                    End If

                End If
            Next
        Next
    End If
    
    If Sx < tx And Sy > ty Then 'Up Right
        For X = Sx + 1 To tx
            For Y = Sy - 1 To ty Step -1
                If Sx - X = Y - Sy Then  'Only check if it's a diagonal
                    If Moving And BoardAI(X, Y) <> 0 Then
                        MoveAlrightAI = False
                        Exit Function
                    End If

                    If Taking And (BoardAI(X, Y) <> 0 And (X <> tx Or Y <> ty)) Then
                        MoveAlrightAI = False
                        Exit Function
                    End If

                End If
            Next
        Next
    End If
    If Sx > tx And Sy < ty Then 'Down Left
        For X = Sx - 1 To tx Step -1
            For Y = Sy + 1 To ty
                If X - Sx = Sy - Y Then 'Only check if it's a diagonal
                    If Moving And BoardAI(X, Y) <> 0 Then
                        MoveAlrightAI = False
                        Exit Function
                    End If

                    If Taking And (BoardAI(X, Y) <> 0 And (X <> tx Or Y <> ty)) Then
                        MoveAlrightAI = False
                        Exit Function
                    End If

                End If
            Next
        Next
    End If
    If Sx < tx And Sy < ty Then 'Down Right
        For X = Sx + 1 To tx Step 1
            For Y = Sy + 1 To ty Step 1
                If Sx - X = Sy - Y Then 'Only check if it's a diagonal
                    If Moving And BoardAI(X, Y) <> 0 Then
                        MoveAlrightAI = False
                        Exit Function
                    End If
                    
                    If Taking And (BoardAI(X, Y) <> 0 And (X <> tx Or Y <> ty)) Then
                        MoveAlrightAI = False
                        Exit Function
                    End If

                End If
            Next
        Next
    End If
    Case 6
        'It's a horsey & they're allowed to jump pieces!
    Case 7
        ' Castling Rule
        If Taking Then
            ' Move illegal so just exit
            MoveAlrightAI = False
            Exit Function
        ElseIf Moving Then
        ' Legal left
        For X = Sx - 1 To tx Step -1
            If Moving And BoardAI(X, Sy) <> 0 Then
                MoveAlrightAI = False
                Exit Function
            End If
        Next X
        
        For X = Sx + 1 To tx Step 1
            If Moving And BoardAI(X, Sy) <> 0 Then
                MoveAlrightAI = False
                Exit Function
            End If
        Next X
        
        End If
    End Select
    
    'Successful Move!
    MoveAlrightAI = True
End If
End Function

Private Function MoveLegalAI(Sx As Integer, Sy As Integer, tx As Integer, ty As Integer) As Boolean
'On Error Resume Next
Dim Legal As Integer
Dim Colour As String
Dim ColourT As String
Dim Adder As Integer
Dim Taking As Boolean
Dim Moving As Boolean
Dim X&, Y&

If BoardAI(Sx, Sy) < 10 Then Colour$ = "B"
If BoardAI(Sx, Sy) > 10 Then Colour$ = "W"
If BoardAI(tx, ty) < 10 And BoardAI(tx, ty) > 0 Then ColourT = "B"
If BoardAI(tx, ty) > 10 Then ColourT = "W"

If Colour = "W" Then Adder = 10

If BoardAI(Sx, Sy) = 0 Or tx > 7 Or ty > 7 Or (Sx = tx And Sy = ty) Or Colour$ = ColourT Then    ' Invalid move
    MoveLegalAI = False
    Exit Function
Else
    Legal = 0

    'See if were trying to take a piece
    If Colour$ <> ColourT And ColourT <> "" Then
        Taking = True
        Moving = False
    Else ' We are just moving to a blank space.
        Taking = False
        Moving = True
    End If
    
    Select Case BoardAI(Sx, Sy)
    Case bPawn + Adder 'Prawn movement
        If Taking Then
            If ty = Sy - 1 And Colour$ = "W" And (tx = Sx - 1 Or tx = Sx + 1) Then Legal = 5  'Black Up + Left/Right to White Taking a piece
            If ty = Sy + 1 And Colour$ = "B" And (tx = Sx - 1 Or tx = Sx + 1) Then Legal = 5  'White Up + Left/Right to Black Taking a piece
        Else
            If (BoardAI(tx, ty) = 0 And ty = Sy - 1 And Colour$ = "W" And Sx = tx) Then Legal = 1    'White vertical
            If (BoardAI(tx, ty) = 0 And ty = Sy + 1 And Colour$ = "B" And Sx = tx) Then Legal = 3    'Black Vertical
            If (Colour$ = "W" And Sy = 7 And ty = Sy - 2) And tx = Sx Then Legal = 1  'First move may be double WHITE
            If (Colour$ = "B" And Sy = 2 And ty = Sy + 2) And tx = Sx Then Legal = 3  'First move may be double BLACK
        End If

    Case bKing + Adder   'King movement
        If tx <= Sx + 1 And tx >= Sx - 1 And ty <= Sy + 1 And ty >= Sy - 1 Then Legal = 5 'Move in any direction by one
        
        ' Allow castling to the right
        If tx < 7 Then _
            If BoardAI(tx + 1, ty) = bRook + Adder And _
                tx = Sx + 2 And ty = Sy Then Legal = 7
            
        ' Allow castling to the Left
        If tx > 2 Then _
            If BoardAI(tx - 2, ty) = bRook + Adder And _
                tx = Sx - 2 And ty = Sy Then Legal = 7
            
    Case bQueen + Adder
        'Diagonal
        If Abs(Sx - tx) = Abs(Sy - ty) Or Abs(tx - Sx) = Abs(ty - Sy) Then Legal = 5
        'Left
        If tx < Sx And ty = Sy Then Legal = 4
        'Right
        If tx > Sx And ty = Sy Then Legal = 2
        'Up
        If tx = Sx And ty < Sy Then Legal = 1
        'Down
        If tx = Sx And ty > Sy Then Legal = 3
        
    Case bBishop + Adder
        If Abs(Sx - tx) = Abs(Sy - ty) Or Abs(tx - Sx) = Abs(ty - Sy) Then Legal = 5
        
    Case bKnight + Adder 'Horsey
        If (tx = Sx + 2 And ty = Sy + 1) Then Legal = 6
        If (tx = Sx + 1 And ty = Sy + 2) Then Legal = 6
        If (tx = Sx - 2 And ty = Sy - 1) Then Legal = 6
        If (tx = Sx - 1 And ty = Sy - 2) Then Legal = 6
        If (tx = Sx + 2 And ty = Sy - 1) Then Legal = 6
        If (tx = Sx + 1 And ty = Sy - 2) Then Legal = 6
        If (tx = Sx - 2 And ty = Sy + 1) Then Legal = 6
        If (tx = Sx - 1 And ty = Sy + 2) Then Legal = 6
    
    Case bRook + Adder 'Castle (Rook)
        'Left
        If tx < Sx And ty = Sy Then Legal = 4
        'Right
        If tx > Sx And ty = Sy Then Legal = 2
        'Up
        If tx = Sx And ty < Sy Then Legal = 1
        'Down
        If tx = Sx And ty > Sy Then Legal = 3
    
    End Select
     'Check to see if we can move, without jumping a piece, based on Legal values
    
    Select Case Legal
    Case 0
        MoveLegalAI = False
        Exit Function
    Case 1 'Up
        For Y = Sy - 1 To ty Step -1
            If Y < 0 Then
                MoveLegalAI = False
                Exit Function
            End If
            
            If Moving And BoardAI(Sx, Y) <> 0 Then
                MoveLegalAI = False
                Exit Function
            End If

            ' are we taking, is the current square blank, if its not blank is it the desired one
            If Taking And (BoardAI(Sx, Y) <> 0 And (Sx <> tx Or Y <> ty)) Then
                MoveLegalAI = False
                Exit Function
            End If
        Next
    Case 2 'Right
        For X = Sx + 1 To tx Step 1
            If X < 0 Then
                MoveLegalAI = False
                Exit Function
            End If

            If Moving And BoardAI(X, Sy) <> 0 Then
                MoveLegalAI = False
                Exit Function
            End If

            If Taking And (BoardAI(X, Sy) <> 0 And (X <> tx Or Sy <> ty)) Then
                MoveLegalAI = False
                Exit Function
            End If

        Next
    Case 3 'Down
        For Y = Sy + 1 To ty Step 1
            If Y < 0 Then
                MoveLegalAI = False
                Exit Function
            End If

            If Moving And BoardAI(Sx, Y) <> 0 Then
                MoveLegalAI = False
                Exit Function
            End If

            If Taking And (BoardAI(Sx, Y) <> 0 And (Sx <> tx Or Y <> ty)) Then
                MoveLegalAI = False
                Exit Function
            End If

        Next
    Case 4 'Left
        For X = Sx - 1 To tx Step -1
            If X < 0 Then
                MoveLegalAI = False
                Exit Function
            End If

            If Moving And BoardAI(X, Sy) <> 0 Then
                MoveLegalAI = False
                Exit Function
            End If

            If Taking And (BoardAI(X, Sy) <> 0 And (X <> tx Or Sy <> ty)) Then
                MoveLegalAI = False
                Exit Function
            End If

        Next
    Case 5 'Diagonal
    If Sx > tx And Sy > ty Then 'Up Left
        For X = Sx - 1 To tx Step -1
            For Y = Sy - 1 To ty Step -1
                If Sx - X = Sy - Y Then 'Only check if it's a diagonal
                    If Moving And BoardAI(X, Y) <> 0 Then
                        MoveLegalAI = False
                        Exit Function
                    End If

                    If Taking And BoardAI(X, Y) <> 0 And (X <> tx Or Y <> ty) Then
                        MoveLegalAI = False
                        Exit Function
                    End If

                End If
            Next
        Next
    End If
    
    If Sx < tx And Sy > ty Then 'Up Right
        For X = Sx + 1 To tx
            For Y = Sy - 1 To ty Step -1
                If Sx - X = Y - Sy Then  'Only check if it's a diagonal
                    If Moving And BoardAI(X, Y) <> 0 Then
                        MoveLegalAI = False
                        Exit Function
                    End If

                    If Taking And (BoardAI(X, Y) <> 0 And (X <> tx Or Y <> ty)) Then
                        MoveLegalAI = False
                        Exit Function
                    End If

                End If
            Next
        Next
    End If
    If Sx > tx And Sy < ty Then 'Down Left
        For X = Sx - 1 To tx Step -1
            For Y = Sy + 1 To ty
                If X - Sx = Sy - Y Then 'Only check if it's a diagonal
                    If Moving And BoardAI(X, Y) <> 0 Then
                        MoveLegalAI = False
                        Exit Function
                    End If

                    If Taking And (BoardAI(X, Y) <> 0 And (X <> tx Or Y <> ty)) Then
                        MoveLegalAI = False
                        Exit Function
                    End If

                End If
            Next
        Next
    End If
    If Sx < tx And Sy < ty Then 'Down Right
        For X = Sx + 1 To tx Step 1
            For Y = Sy + 1 To ty Step 1
                If Sx - X = Sy - Y Then 'Only check if it's a diagonal
                    If Moving And BoardAI(X, Y) <> 0 Then
                        MoveLegalAI = False
                        Exit Function
                    End If
                    
                    If Taking And (BoardAI(X, Y) <> 0 And (X <> tx Or Y <> ty)) Then
                        MoveLegalAI = False
                        Exit Function
                    End If

                End If
            Next
        Next
    End If
    Case 6
        'It's a horsey & they're allowed to jump pieces!
    Case 7
        ' Castling Rule
        If Taking Then
            ' Move illegal so just exit
            MoveLegalAI = False
            Exit Function
        ElseIf Moving Then
        ' Legal left
        For X = Sx - 1 To tx Step -1
            If Moving And BoardAI(X, Sy) <> 0 Then
                MoveLegalAI = False
                Exit Function
            End If
        Next X
        
        For X = Sx + 1 To tx Step 1
            If Moving And BoardAI(X, Sy) <> 0 Then
                MoveLegalAI = False
                Exit Function
            End If
        Next X
        
        End If
    End Select
    
    'Successful Move!
    MoveLegalAI = True
End If
End Function

Private Sub RenderBoard()
Dim Piece As Integer, Mask As Integer
Dim X As Integer, Y As Integer

picFrontBuffer.Cls

For X = 0 To 7
For Y = 0 To 7
If Board(X, Y) Then
Piece = Board(X, Y)
Mask = Piece + 10
If Mask < 20 Then Mask = Mask + 10
If mnuDangerShow.Checked And CheckPosition(X, Y) Then Call AlphaBlending(picFrontBuffer.hdc, 48 * X, 48 * Y, 48, 48, picEffector(1).hdc, 0, 0, 48, 48, &H800000)
BitBlt picFrontBuffer.hdc, 48 * X, 48 * Y, 48, 48, picPiece(Mask).hdc, 0, 0, vbMergePaint
BitBlt picFrontBuffer.hdc, 48 * X, 48 * Y, 48, 48, picPiece(Piece).hdc, 0, 0, vbSrcAnd
If Selected And SelX = X And SelY = Y Then Call AlphaBlending(picFrontBuffer.hdc, 48 * X, 48 * Y, 48, 48, picEffector(0).hdc, 0, 0, 48, 48, &H800000)
End If
Next Y
Next X

picFrontBuffer.Refresh
DoEvents
End Sub

Private Sub RenderBoardAI()
Dim Piece As Integer, Mask As Integer
Dim X As Integer, Y As Integer

picFrontBuffer.Cls

For X = 0 To 7
For Y = 0 To 7
If BoardAI(X, Y) Then
Piece = BoardAI(X, Y)
Mask = Piece + 10
If Mask < 20 Then Mask = Mask + 10
If mnuDangerShow.Checked And CheckPositionAI(X, Y) Then Call AlphaBlending(picFrontBuffer.hdc, 48 * X, 48 * Y, 48, 48, picEffector(1).hdc, 0, 0, 48, 48, &H800000)
BitBlt picFrontBuffer.hdc, 48 * X, 48 * Y, 48, 48, picPiece(Mask).hdc, 0, 0, vbMergePaint
BitBlt picFrontBuffer.hdc, 48 * X, 48 * Y, 48, 48, picPiece(Piece).hdc, 0, 0, vbSrcAnd
End If
Next Y
Next X

picFrontBuffer.Refresh
DoEvents
End Sub

Function CheckPosition(KingX%, KingY%, Optional ReplacePiece As Byte = 0) As Boolean
'On Error Resume Next

Dim X As Integer, Y As Integer, OldTurn As Integer
Dim OldPiece As Byte

CheckPosition = False

If Board(KingX, KingY) = 0 Or ReplacePiece = 15 Or ReplacePiece = 5 Then OldPiece = Board(KingX, KingY): Board(KingX, KingY) = ReplacePiece

For X = 0 To 7 'See if the piece is checked
For Y = 0 To 7
If (Board(X, Y) < 10 And Board(KingX, KingY) > 10) Or (Board(X, Y) > 10 And Board(KingX, KingY) < 10) Then
If MoveLegal(X, Y, KingX, KingY) Then
CheckPosition = True 'CHECK
If Board(KingX, KingY) = ReplacePiece Then Board(KingX, KingY) = OldPiece
Exit Function
End If
End If
Next
Next

If Board(KingX, KingY) = ReplacePiece Then Board(KingX, KingY) = OldPiece

'If Err Then MsgBox Err.Description, vbCritical, "Serious Error"
End Function

Function CheckPositionAI(KingX%, KingY%, Optional ReplacePiece As Byte = 0) As Boolean
On Error Resume Next

Dim X As Integer, Y As Integer, OldTurn As Integer
Dim OldPiece As Byte

CheckPositionAI = False

If Board(KingX, KingY) = 0 Or ReplacePiece = 15 Or ReplacePiece = 5 Then OldPiece = Board(KingX, KingY): Board(KingX, KingY) = ReplacePiece

For X = 0 To 7 'See if the piece is checked
For Y = 0 To 7
If (Board(X, Y) < 10 And Turn = 0) Or (Board(X, Y) > 10 And Turn = 1) Then
If MoveLegal(X, Y, KingX, KingY) Then
CheckPositionAI = True 'CHECK
If Board(KingX, KingY) = ReplacePiece Then Board(KingX, KingY) = OldPiece
Exit Function
End If
End If
Next
Next

If Board(KingX, KingY) = ReplacePiece Then Board(KingX, KingY) = OldPiece

If Err Then MsgBox Err.Description, vbCritical, "Serious Error"
End Function

Private Sub Label1_Click()
If Mode <> "" Then Exit Sub
PopupMenu mnuGametype, , 232, 96
End Sub

Private Sub Label2_Click()
If Mode <> "" Then Exit Sub
PopupMenu mnuAIType, , 232, 112
End Sub

Private Sub mnuAbout_Click()
frmAbout.Show
End Sub

Private Sub mnuAI_Click(Index As Integer)
mnuAI(0).Checked = False
mnuAI(1).Checked = False
mnuAI(2).Checked = False
mnuAI(3).Checked = False
mnuAI(Index).Checked = True
End Sub

Private Sub mnuDangerShow_Click()
If mnuDangerShow.Checked Then mnuDangerShow.Checked = False Else mnuDangerShow.Checked = True
End Sub

Private Sub mnuDebugVerbose_Click()
If mnuDebugVerbose.Checked Then mnuDebugVerbose.Checked = False Else mnuDebugVerbose.Checked = True
End Sub

Private Sub mnuExit_Click()
Unload Me
End Sub

Private Sub mnuGamePlayers_Click(Index As Integer)
mnuGamePlayers(0).Checked = False
mnuGamePlayers(1).Checked = False
mnuGamePlayers(2).Checked = False
mnuGamePlayers(Index).Checked = True
End Sub

Private Sub mnuViewLog_Click()
frmLogViewer.Show
End Sub

Private Sub picFrontBuffer_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
If Button <> vbLeftButton Then Exit Sub

Lx = Int(X / 48)
Ly = Int(Y / 48)

If Lx > 7 Then Lx = 7
If Ly > 7 Then Ly = 7
If Lx < 0 Then Lx = 0
If Ly < 0 Then Ly = 0

If Selected = True And Lx = SelX And Ly = SelY Then
Selected = False
Call RenderBoard
Exit Sub
End If

If Turn = 0 And Board(Lx, Ly) <> 0 And Board(Lx, Ly) < 10 And Selected = False Then
SelX = Lx
SelY = Ly
Selected = True
Call RenderBoard
Exit Sub
End If

If Turn = 1 And Board(Lx, Ly) > 10 And Selected = False Then
SelX = Lx
SelY = Ly
Selected = True
Call RenderBoard
Exit Sub
End If

If Selected = True And (Lx <> SelX Or Ly <> SelY) Then
Call MakeMove(SelX, SelY, Lx, Ly)
Selected = False
Call RenderBoard
Exit Sub
End If

Call RenderBoard
End Sub

Private Sub picFrontBuffer_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Lx = Int(X / 48)
Ly = Int(Y / 48)

If Lx > 7 Then Lx = 7
If Ly > 7 Then Ly = 7
If Lx < 0 Then Lx = 0
If Ly < 0 Then Ly = 0

Call StatusReport
End Sub

Private Sub StatusReport()
Dim Status As String, XCoord$, XCoord2$, PieceName$, PieceName2$

'OfflinePP , OfflinePC, OfflineCC, Client, Server
'Status = "ChessMASTER Mark III" + vbNewLine

Select Case Mode
Case "OfflinePP"
Status = Status + "Offline Two Player Mode" + vbNewLine
Case "OfflinePC"
Status = Status + "Offline Playing CPU Mode" + vbNewLine
Case "OfflineCC"
Status = Status + "CPU vs CPU Mode" + vbNewLine
Case "Client"
Status = Status + "Online Client Mode" + vbNewLine
Case "Server"
Status = Status + "Online Server Mode" + vbNewLine
Case "EditMode"
Status = Status + "Board Editing Mode" + vbNewLine
Case ""
Status = Status + "ChessMASTER Mark III" + vbNewLine
End Select

Select Case Turn
Case 0
If Mode = "OfflinePP" Then Status = Status + "Waiting for Black"
If Mode = "OfflinePC" Then Status = Status + "Waiting for CPU"
If Mode = "OfflineCC" Then Status = Status + "Waiting for Black"
'If Mode = "Client" Then Status = Status + "Your Turn"
'If Mode = "Server" Then Status = Status + "Waiting for" + WsServe.RemoteHost
Case 1
If Mode = "OfflinePP" Then Status = Status + "Waiting for White"
If Mode = "OfflinePC" Then Status = Status + "Waiting for White"
If Mode = "OfflineCC" Then Status = Status + "Waiting for White"
'If Mode = "Client" Then Status = Status + "Waiting for" + WsServe.RemoteHost
'If Mode = "Server" Then Status = Status + "Your Turn"
End Select

Select Case SelX
Case 0: XCoord$ = "A"
Case 1: XCoord$ = "B"
Case 2: XCoord$ = "C"
Case 3: XCoord$ = "D"
Case 4: XCoord$ = "E"
Case 5: XCoord$ = "F"
Case 6: XCoord$ = "G"
Case 7: XCoord$ = "H"
End Select

Select Case Lx
Case 0: XCoord2$ = "A"
Case 1: XCoord2$ = "B"
Case 2: XCoord2$ = "C"
Case 3: XCoord2$ = "D"
Case 4: XCoord2$ = "E"
Case 5: XCoord2$ = "F"
Case 6: XCoord2$ = "G"
Case 7: XCoord2$ = "H"
End Select

Select Case Board(SelX, SelY)
Case 0: PieceName = "Empty"
Case bRook: PieceName = "Rook"
Case bKnight: PieceName = "Knight"
Case bBishop: PieceName = "Bishop"
Case bQueen: PieceName = "Queen"
Case bKing: PieceName = "King"
Case bPawn: PieceName = "Pawn"
Case wRook: PieceName = "Rook"
Case wKnight: PieceName = "Knight"
Case wBishop: PieceName = "Bishop"
Case wQueen: PieceName = "Queen"
Case wKing: PieceName = "King"
Case wPawn: PieceName = "Pawn"
End Select

Select Case Board(Lx, Ly)
Case 0: PieceName2 = "Empty"
Case bRook: PieceName2 = "Rook"
Case bKnight: PieceName2 = "Knight"
Case bBishop: PieceName2 = "Bishop"
Case bQueen: PieceName2 = "Queen"
Case bKing: PieceName2 = "King"
Case bPawn: PieceName2 = "Pawn"
Case wRook: PieceName2 = "Rook"
Case wKnight: PieceName2 = "Knight"
Case wBishop: PieceName2 = "Bishop"
Case wQueen: PieceName2 = "Queen"
Case wKing: PieceName2 = "King"
Case wPawn: PieceName2 = "Pawn"
End Select

If Selected = True Then
Status = Status + vbNewLine + PieceName + " at " + XCoord + ":" + Trim$(Str$(SelY + 1)) + " To " + XCoord2 + ":" + Trim$(Str$(Ly + 1)) + " (" + PieceName2 + ")"
picFrontBuffer.ToolTipText = PieceName + " at " + XCoord + ":" + Trim$(Str$(SelY + 1)) + " To " + XCoord2 + ":" + Trim$(Str$(Ly + 1)) + " (" + PieceName2 + ")"
If MoveAlright(SelX, SelY, Lx, Ly) Then Status = Status + vbNewLine + "~Move Legal~": picFrontBuffer.ToolTipText = "[Legal] " + picFrontBuffer.ToolTipText
Else
Status = Status + vbNewLine + PieceName2 + " at " + XCoord2 + ":" + Trim$(Str(Ly + 1))
If Board(Lx, Ly) < 10 Then picFrontBuffer.ToolTipText = "Black " + PieceName2 + " at " + XCoord2 + ":" + Trim$(Str(Ly + 1))
If Board(Lx, Ly) > 10 Then picFrontBuffer.ToolTipText = "White " + PieceName2 + " at " + XCoord2 + ":" + Trim$(Str(Ly + 1))
If Board(Lx, Ly) = 0 Then picFrontBuffer.ToolTipText = PieceName2 + " at " + XCoord2 + ":" + Trim$(Str(Ly + 1))
End If

lblStatus.Caption = Status$
End Sub

Private Sub PerformAI()
Randomize GetTickCount

Label2.ForeColor = vbYellow
Label2.Refresh

Select Case AIType
Case 0: Call WoodPush 'WoodPusher!
Case 1: Call AgreeAI 'ChessMASTER Mark I AI
Case 2: Call ClassicAI 'ChessMASTER Mark II AI
Case 3: Call BestAI 'ChessMASTER Mark III AI
End Select

Label2.ForeColor = &HC0&
Label2.Refresh

Timer1.Enabled = True
End Sub

Private Sub WoodPush()
Dim StartX(4096) As Byte, StartY(4096) As Byte
Dim EndX(4096) As Byte, EndY(4096) As Byte
Dim StartTime As Long, EndTime As Long
Dim CurrentIndex As Integer
Dim X As Integer, Y As Integer
Dim i As Integer, J As Integer

StartTime = GetTickCount

If Turn = 0 Then 'WoodPusher is Black
For X = 0 To 7
For Y = 0 To 7
If Board(X, Y) < 10 Then 'Get Peice
For i = 0 To 7
For J = 0 To 7
If Not (X = i And J = Y) Then
If MoveLegal(X, Y, i, J) Then 'Move Legal, So log
StartX(CurrentIndex) = X
StartY(CurrentIndex) = Y
EndX(CurrentIndex) = i
EndY(CurrentIndex) = J
CurrentIndex = CurrentIndex + 1
End If
End If
Next J
Next i
End If
Next Y
Next X
End If

If Turn = 1 Then 'WoodPusher is White
For X = 0 To 7
For Y = 0 To 7
If Board(X, Y) > 10 Then 'Get Peice
For i = 0 To 7
For J = 0 To 7
If Not (X = i And J = Y) Then
If MoveLegal(X, Y, i, J) Then 'Move Legal, So log
StartX(CurrentIndex) = X
StartY(CurrentIndex) = Y
EndX(CurrentIndex) = i
EndY(CurrentIndex) = J
CurrentIndex = CurrentIndex + 1
End If
End If
Next J
Next i
End If
Next Y
Next X
End If

Let i = Int(Rnd * CurrentIndex) 'Choose Random Move
Call MakeMove(CInt(StartX(i)), CInt(StartY(i)), CInt(EndX(i)), CInt(EndY(i)))

EndTime = GetTickCount

If mnuDebugVerbose.Checked Then Call MsgBox(" NSE ChessMASTER Series, WoodPusher AI." + vbNewLine + Str$(CurrentIndex) + " Moves calculated and move #" + Str$(i) + " of it chosen at random in" + Str$((EndTime - StartTime) / 1000) + " seconds.", vbInformation, "NSE ChessMASTER Mark III")
End Sub

Private Sub AgreeAI() 'ChessMASTER Mark I AI
Dim X As Integer, Y As Integer
Dim X2 As Integer, Y2 As Integer
Dim StartTime As Long, EndTime As Long
Dim HypFrX%(72)
Dim HypFrY%(72)
Dim HypToX%(72)
Dim HypToY%(72)
Dim HypQuality(72)
Dim NumberOfBest As Integer, BestSoFar As Integer
Dim i As Long

StartTime = GetTickCount

'AI. This AI uses the actual board memory & therfor cannot be recursive above 1 level.
'If it used a copy of the board it could minipulate and play ahead (?recusivly).

'For X = 0 To 7
'For Y = 0 To 7
'BoardAI(X, Y) = Board(X, Y)
'Next Y
'Next X

BoardAI() = Board()

'Scan board and move all possibilities giving a score for taking enemy peices

If Turn = 0 Then
For X = 0 To 7
For Y = 0 To 7
If BoardAI(X, Y) < 10 And BoardAI(X, Y) <> 0 Then
For X2 = 0 To 7
For Y2 = 0 To 7
If MoveLegalAI(X, Y, X2, Y2) Then
HypFrX%(X2 * 8 + Y2) = X
HypFrY%(X2 * 8 + Y2) = Y
HypToX%(X2 * 8 + Y2) = X2
HypToY%(X2 * 8 + Y2) = Y2
HypQuality(X2 * 8 + Y2) = PieceValue(True, X2, Y2)
End If
Next
Next
End If
Next
Next
'Scan board moving all possibities of enemies and subtracting score from move locations
For X = 0 To 7
For Y = 0 To 7
If BoardAI(X, Y) < 10 And BoardAI(X, Y) <> 0 Then
For X2 = 0 To 7
For Y2 = 0 To 7
If MoveLegalAI(X, Y, X2, Y2) Then
HypQuality(X2 * 8 + Y2) = HypQuality(X2 * 8 + Y2) - PieceValue(True, X, Y)
End If
Next
Next
End If
Next
Next
End If

If Turn = 1 Then
For X = 0 To 7
For Y = 0 To 7
If BoardAI(X, Y) > 10 Then
For X2 = 0 To 7
For Y2 = 0 To 7
If MoveLegalAI(X, Y, X2, Y2) Then
HypFrX%(X2 * 8 + Y2) = X
HypFrY%(X2 * 8 + Y2) = Y
HypToX%(X2 * 8 + Y2) = X2
HypToY%(X2 * 8 + Y2) = Y2
HypQuality(X2 * 8 + Y2) = PieceValue(True, X2, Y2)
End If
Next
Next
End If
Next
Next
'Scan board moving all possibities of enemies and subtracting score from move locations
For X = 0 To 7
For Y = 0 To 7
If BoardAI(X, Y) > 10 Then
For X2 = 0 To 7
For Y2 = 0 To 7
If MoveLegalAI(X, Y, X2, Y2) Then
HypQuality(X2 * 8 + Y2) = HypQuality(X2 * 8 + Y2) - PieceValue(True, X, Y)
End If
Next
Next
End If
Next
Next
End If

'Scan accumulated score & coordinate arrays & execute the one with the greatest score
NumberOfBest = 0
BestSoFar = -100
For i = 0 To 72
If MoveAlright(HypFrX(i), HypFrY(i), HypToX(i), HypToY(i)) And Not HypQuality(i) = Empty And HypQuality(i) > BestSoFar Or (HypQuality(i) = BestSoFar And Rnd < 0.2) Then
NumberOfBest = i
BestSoFar = HypQuality(i)
End If
Next

'MsgBox "Invalid Move Produced, Advanced Rule Infringed by Quick checking"

EndTime = GetTickCount

If mnuDebugVerbose.Checked Then Call MsgBox(" NSE ChessMASTER Series, Agree AI." + vbNewLine + " [72 Constant]" + " Moves calculated and move #" + Str$(NumberOfBest) + " of it chosen at random in" + Str$((EndTime - StartTime) / 1000) + " seconds.", vbInformation, "NSE ChessMASTER Mark III")

Call MakeMove(HypFrX(NumberOfBest), HypFrY(NumberOfBest), HypToX(NumberOfBest), HypToY(NumberOfBest))
End Sub

Private Sub ClassicAI() 'ChessMASTER 2 ported AI
Dim StartTime As Long
Dim EndTime As Long
StartTime = GetTickCount

lblStatus.Caption = "ChessMASTER thinking..."
lblStatus.Refresh
Randomize Timer

Dim X As Integer, X2 As Integer
Dim Y As Integer, Y2 As Integer
Dim n As Integer
Dim m As Integer
Dim Opposite As String
Dim BestScore As Integer
Dim HypFrX%(72)
Dim HypFrY%(72)
Dim HypToX%(72)
Dim HypToY%(72)
Dim BestMove As Long
Dim Cost As Integer
Dim i As Integer
BestScore = -32000

Dim pX As Integer
Dim pY As Integer

Dim Score As Integer
Dim MoveIndex As Integer
' Debug info
'Open "C:\Chess.log" For Output As #1
MoveIndex = 0
For X = 0 To 7
    For Y = 0 To 7
        ' If this is our piece then see where we can move to
        If (Board(X, Y) > 10 And Turn = 1) Or (Board(X, Y) < 10 And Turn = 0) Then
            For X2 = 0 To 7
                For Y2 = 0 To 7
                    If MoveLegal(X, Y, X2, Y2) Then
                        ' Get the + score of this move (value of the piece taken)
                            Score = PieceValue(False, X2, Y2)
                        ' Subsidise moving out of a threatened space
                            If CheckPosition(X, Y) And CheckPosition(X2, Y2) = False Then Score = Score + PieceValue(False, X, Y)
                        ' Subtract from subsidy if covered by an inferior peice
                            If CoveredPosition(False, X, Y) Then Score = Score - 0.67 * PieceValue(False, X, Y)

                            
                            ' Reset the temporary board
                            'For pX = 0 To 7
                            '    For pY = 0 To 7
                            '        BoardAI(pX, pY) = Board(pX, pY)
                            '    Next pY
                            'Next pX
                            
                            BoardAI() = Board()
                            
                            ' Move from x -> n and y -> m
                            Dim HypMovedPiece As Byte
                            
                            BoardAI(X2, Y2) = Board(X, Y)
                            BoardAI(X, Y) = 0
                            
                            ' Get the - score of this move (the highest piece we are now exposing to be taken)
                            Cost = moveRisk()
                            
                            Dim DontMoveTheKing As Integer
                            HypMovedPiece = Board(X, Y)
                            ' Don't move the king unless you need to.
                            If (HypMovedPiece = bKing Or HypMovedPiece = wKing) And Board(X2, Y2) = 0 Then
                                DontMoveTheKing = 0
                            Else
                                DontMoveTheKing = 0
                            End If
                            
                            If (Score - (Cost + DontMoveTheKing)) > BestScore Then
'                                Write #1, "-----------------------------------------"
                                BestScore = (Score - (Cost + DontMoveTheKing))
                                
'                                Write #1, "New Best Score Move:" & MovingObject & ">" & Score & "," & Cost
                                HypFrX%(0) = X
                                HypFrY%(0) = Y
                                HypToX%(0) = X2
                                HypToY%(0) = Y2
                                
                                MoveIndex = 0
                                

 '                               Write #1, BestScore & "," & Score & "," & cost & "," & x & "," & y & "," & X2 & "," & Y2 & ","
                            ElseIf (Score - (Cost + DontMoveTheKing)) = BestScore Then
                                MoveIndex = MoveIndex + 1
                                
'                                Write #1, "Additional Best Score Move:" & MovingObject & ">" & Score & "," & Cost
                                HypFrX%(MoveIndex) = X
                                HypFrY%(MoveIndex) = Y
                                HypToX%(MoveIndex) = X2
                                HypToY%(MoveIndex) = Y2

                                
                            
                            End If
                            If mnuAIMultiTask.Checked = True Then DoEvents
                        'End If
                    End If
                Next Y2
            Next X2
        End If
    Next Y
Next X


' Pick Working Move if Best Move is illegal
Do
BestMove = Int(Rnd * MoveIndex)
Loop Until MoveAlright(HypFrX(BestMove), HypFrY(BestMove), HypToX(BestMove), HypToY(BestMove))


lblStatus.Caption = ""

EndTime = GetTickCount

'If mnuDebugVerbose.Checked = True Then Call MsgBox(" NSE ChessMASTER Series, ChessMASTER Mark II AI." + vbNewLine + Str$(moveIndex) + " Moves calculated and move #" + Str$(BestMove) + " of it chosen at random in" + Str$((EndTime - StartTime) / 1000) + " seconds.", vbInformation, "NSE ChessMASTER Mark III")

Call MakeMove(CInt(HypFrX%(BestMove)), CInt(HypFrY%(BestMove)), CInt(HypToX%(BestMove)), CInt(HypToY%(BestMove)))

End Sub

Private Function PieceValue(Hyperthetical As Boolean, Rank As Integer, File As Integer)
Dim Piece As Byte

If Hyperthetical Then Piece = BoardAI(Rank, File) Else Piece = Board(Rank, File)

If Piece > 10 Then Piece = Piece - 10

Select Case Piece
Case 0: PieceValue = -10
Case bRook: PieceValue = 60
Case bKnight: PieceValue = 70
Case bBishop: PieceValue = 60
Case bQueen: PieceValue = 128
Case bKing: PieceValue = 2550
Case bPawn: PieceValue = 20
End Select
End Function

Function CoveredPosition(Hyperthetical As Boolean, Rank As Integer, File As Integer)
Dim X As Integer
Dim Y As Integer

For X = 0 To 7
For Y = 0 To 7
If Hyperthetical = True Then

If (BoardAI(X, Y) < 10 And BoardAI(Rank, File) < 10) Or (BoardAI(X, Y) > 10 And BoardAI(Rank, File) > 10) Then 'If piece is on same team
    If MoveLegalAI(X, Y, Rank, File) Then 'If peice is covered by other peice
        CoveredPosition = True
        Exit Function
    End If
End If

ElseIf Hyperthetical = False Then

If (Board(X, Y) < 10 And Board(Rank, File) < 10) Or (Board(X, Y) > 10 And Board(Rank, File) > 10) Then 'If piece is on same team
    If MoveLegal(X, Y, Rank, File) Then 'If peice is covered by other peice
        CoveredPosition = True
        Exit Function
    End If
End If

End If
Next
Next
End Function

Function moveRisk() As Integer
Dim BestScore As Integer
Dim Score As Integer
Dim X As Integer, X2 As Integer
Dim Y As Integer, Y2 As Integer
Dim n As Integer
Dim m As Integer
BestScore = -100

For X = 0 To 7
    For Y = 0 To 7
        ' IF its an opponents piece then
        If (Turn = 1 And BoardAI(X, Y) < 10) Or (Turn = 0 And BoardAI(X, Y) > 10) Then
            For X2 = 0 To 7
                For Y2 = 0 To 7
                    If MoveLegalAI(X, Y, X2, Y2) Then
                    'Your in Check.
                        
                        Score = PieceValue(True, X2, Y2)
                        If Score > BestScore Then BestScore = Score
'                        Write #1, x, y, X2, Y2, tempBoard(x, y), "moveRisk:" & Score
                  
                    End If
                Next Y2
            Next X2
        End If
    Next Y
Next X

moveRisk = BestScore
End Function

Private Sub BestAI() 'Ni-Star Enterprises 24/12/2002
ReDim StartX(4096) As Byte, StartY(4096) As Byte
ReDim EndX(4096) As Byte, EndY(4096) As Byte
ReDim MoveScore(4096) As Integer
Dim WorkBoard() As Byte
Dim X As Byte, Y As Byte
Dim X2 As Byte, Y2 As Byte
Dim i As Integer
Dim MoveIndex As Integer
Dim ProScore As Integer, ConScore As Integer

Dim BestMoveScore As Integer
Dim BestMoveIndex As Integer

MoveIndex = 0

WorkBoard() = Board()
BoardAI() = WorkBoard()

For X = 0 To 7 'Find All Legal Moves
For Y = 0 To 7
        For X2 = 0 To 7
        For Y2 = 0 To 7
            If (Turn = 0 And BoardAI(X, Y) < 10) Or (Turn = 1 And BoardAI(X, Y) > 10) Then
                If MoveLegalAI(CInt(X), CInt(Y), CInt(X2), CInt(Y2)) Then
                    StartX(MoveIndex) = X
                    StartY(MoveIndex) = Y
                    EndX(MoveIndex) = X2
                    EndY(MoveIndex) = Y2
                    MoveIndex = MoveIndex + 1
                End If
            End If
        Next Y2
        Next X2
Next Y
Next X

MoveIndex = MoveIndex - 1

ReDim Preserve StartX(MoveIndex) As Byte, StartY(MoveIndex) As Byte
ReDim Preserve EndX(MoveIndex) As Byte, EndY(MoveIndex) As Byte
ReDim Preserve MoveScore(MoveIndex) As Integer

For i = 0 To MoveIndex
'Value of taken Piece +
ProScore = PieceValue(True, CInt(EndX(i)), CInt(EndY(i)))
'If Indanger Add ProScore to Move it
If CheckPositionAI(CInt(StartX(i)), CInt(StartY(i))) Then ProScore = ProScore + PieceValue(True, CInt(StartX(i)), CInt(StartY(i)))
'Value of Piece Risked -
BoardAI(EndX(i), EndY(i)) = BoardAI(StartX(i), StartY(i))
If CheckPositionAI(CInt(EndX(i)), CInt(EndY(i))) Then ConScore = PieceValue(True, CInt(EndX(i)), CInt(EndY(i)))

BoardAI() = WorkBoard()

MoveScore(i) = ProScore - ConScore
Next i


For i = 0 To MoveIndex 'Find Best Score
If MoveScore(i) > MoveScore(BestMoveIndex) Then
If MoveAlright(CInt(StartX(BestMoveIndex)), CInt(StartY(BestMoveIndex)), CInt(EndX(BestMoveIndex)), CInt(EndY(BestMoveIndex))) Then BestMoveIndex = i
End If
Next i

For i = 0 To MoveIndex 'Mark Out Lower Scores
If MoveScore(i) < MoveScore(BestMoveIndex) Then
MoveScore(i) = -32766
End If
Next i

Do 'Choose Randomly one of highest scores
BestMoveIndex = CInt(Rnd * MoveIndex)
Loop Until MoveScore(BestMoveIndex) > -32766

MakeMove CInt(StartX(BestMoveIndex)), CInt(StartY(BestMoveIndex)), CInt(EndX(BestMoveIndex)), CInt(EndY(BestMoveIndex))

'MsgBox Str$(MoveIndex), vbSystemModal + vbExclamation, "NSE ChessMASTER BESTAI"

End Sub

Private Sub EndGame()
Dim X As Integer
Dim Y As Integer

Mode = ""
Turn = -1

For X = 0 To 7
For Y = 0 To 7
Board(X, Y) = 0
Next Y
Next X

ReDim BoardAI(7, 7) As Byte

Call RenderBoard
End Sub

Private Function CheckForCheck()
Dim Adder As Integer
Dim X As Integer, Y As Integer
Dim i As Integer, J As Integer

If Turn = 1 Then Adder = 10

For X = 0 To 7
For Y = 0 To 7
If Board(X, Y) = bKing + Adder Then
For i = 0 To 7
For J = 0 To 7
If (Board(i, J) < 10 And Turn = 1) Or (Board(i, J) > 10 And Turn = 0) Then
If MoveLegal(i, J, X, Y) Then
CheckForCheck = True
Exit Function
End If
End If
Next J
Next i
End If
Next Y
Next X
End Function

Private Sub CheckGame() 'Check mate?
Dim wCheckMate As Boolean, bCheckMate As Boolean
Dim wCheck As Boolean, bCheck As Boolean
Dim bKingX As Integer, bKingY As Integer
Dim wKingX As Integer, wKingY As Integer
Dim Threats As Long
Dim X As Integer, Y As Integer
Dim Xi As Integer, Yi As Integer
Dim Rematch As Integer

wKingX = -32766: wKingY = -32766
bKingX = -32766: bKingY = -32766
For X = 0 To 7
For Y = 0 To 7
If Board(X, Y) = wKing Then wKingX = X: wKingY = Y
If Board(X, Y) = bKing Then bKingX = X: bKingY = Y
Next Y
Next X

If (wKingX = -32766 And wKingY = -32766) Then
Rematch = MsgBox("Black Win. White Lose. Play again?", vbYesNo + vbQuestion, "NSE ChessMASTER Mark III")
If Rematch = vbYes Then
InitiateNewGame
Exit Sub
Else
EndGame
Exit Sub
End If
End If
If (bKingX = -32766 And bKingY = -32766) Then
Rematch = MsgBox("White Win. Black Lose. Play again?", vbYesNo + vbQuestion, "NSE ChessMASTER Mark III")
If Rematch = vbYes Then
InitiateNewGame
Exit Sub
Else
EndGame
Exit Sub
End If
End If

wCheck = CheckPosition(wKingX, wKingY)
If wCheck = False Then 'Check if White are checkmated
wCheckMate = False
Else
wCheckMate = True
For Xi = wKingX - 1 To wKingX + 1 Step 1
For Yi = wKingY - 1 To wKingY + 1 Step 1
If Xi >= 0 And Xi < 8 And Yi >= 0 And Yi < 8 Then
If (wCheck And Board(Xi, Yi) = 0) Or (wCheck = False And Board(Xi, Yi) = wKing) Or Board(Xi, Yi) < 10 Then
Threats = 0
For X = 0 To 7 Step 1
For Y = 0 To 7 Step 1
If Board(X, Y) < 10 And X <> Xi And Y <> Yi Then
If MoveAlright(X, Y, Xi, Yi) Then
Threats = Threats + 1
End If
End If
Next Y
Next X
If Threats = 0 Then wCheckMate = False: Exit For
End If
End If
If wCheckMate = False Then Exit For
Next Yi
If wCheckMate = False Then Exit For
Next Xi
End If

If wCheckMate Then
'MsgBox "White are in checkmate"
Rematch = MsgBox("Black Win. White Lose. Play again?", vbYesNo + vbQuestion, "NSE ChessMASTER Mark III")
If Rematch = vbYes Then
InitiateNewGame
Exit Sub
Else
EndGame
Exit Sub
End If
End If
bCheck = CheckPosition(bKingX, bKingY)
If bCheck = False Then 'Check if White are checkmated
bCheckMate = False
Else
bCheckMate = True
For Xi = bKingX - 1 To bKingX + 1 Step 1
For Yi = bKingY - 1 To bKingY + 1 Step 1
If Xi >= 0 And Xi < 8 And Yi >= 0 And Yi < 8 Then
If (wCheck And Board(Xi, Yi) = 0) Or (wCheck = False And Board(Xi, Yi) = bKing) Or Board(Xi, Yi) > 10 Then
Threats = 0
For X = 0 To 7 Step 1
For Y = 0 To 7 Step 1
If Board(X, Y) > 10 And X <> Xi And Y <> Yi Then
If MoveAlright(X, Y, Xi, Yi) Then
Threats = Threats + 1
End If
End If
Next Y
Next X
If Threats = 0 Then bCheckMate = False: Exit For
End If
End If
If bCheckMate = False Then Exit For
Next Yi
If bCheckMate = False Then Exit For
Next Xi
End If

If bCheckMate Then
'MsgBox "Black are in checkmate"
Rematch = MsgBox("White Win. Black Lose. Play again?", vbYesNo + vbQuestion, "NSE ChessMASTER Mark III")
If Rematch = vbYes Then
InitiateNewGame
Exit Sub
Else
EndGame
Exit Sub
End If
End If

DoEvents
End Sub

Public Function MoveLogReturn() As String
MoveLogReturn = MakeMoveRecord
End Function

Private Sub LogMove(ByRef Log As String, Rank1 As Integer, File1 As Integer, Rank2 As Integer, File2 As Integer)
Dim Status As String, XCoord$, XCoord2$, PieceName$, PieceName2$

Select Case Turn
Case 0
If Mode = "OfflinePP" Then Status = Status + "Waiting for Black"
If Mode = "OfflinePC" Then Status = Status + "Waiting for CPU"
If Mode = "OfflineCC" Then Status = Status + "Waiting for Black"
'If Mode = "Client" Then Status = Status + "Your Turn"
'If Mode = "Server" Then Status = Status + "Waiting for" + WsServe.RemoteHost
Case 1
If Mode = "OfflinePP" Then Status = Status + "Waiting for White"
If Mode = "OfflinePC" Then Status = Status + "Waiting for White"
If Mode = "OfflineCC" Then Status = Status + "Waiting for White"
'If Mode = "Client" Then Status = Status + "Waiting for" + WsServe.RemoteHost
'If Mode = "Server" Then Status = Status + "Your Turn"
End Select

Select Case Rank1
Case 0: XCoord$ = "A"
Case 1: XCoord$ = "B"
Case 2: XCoord$ = "C"
Case 3: XCoord$ = "D"
Case 4: XCoord$ = "E"
Case 5: XCoord$ = "F"
Case 6: XCoord$ = "G"
Case 7: XCoord$ = "H"
End Select

Select Case Rank2
Case 0: XCoord2$ = "A"
Case 1: XCoord2$ = "B"
Case 2: XCoord2$ = "C"
Case 3: XCoord2$ = "D"
Case 4: XCoord2$ = "E"
Case 5: XCoord2$ = "F"
Case 6: XCoord2$ = "G"
Case 7: XCoord2$ = "H"
End Select

Select Case Board(Rank1, File1)
Case 0: PieceName = "Empty"
Case bRook: PieceName = "Rook"
Case bKnight: PieceName = "Knight"
Case bBishop: PieceName = "Bishop"
Case bQueen: PieceName = "Queen"
Case bKing: PieceName = "King"
Case bPawn: PieceName = "Pawn"
Case wRook: PieceName = "Rook"
Case wKnight: PieceName = "Knight"
Case wBishop: PieceName = "Bishop"
Case wQueen: PieceName = "Queen"
Case wKing: PieceName = "King"
Case wPawn: PieceName = "Pawn"
End Select

Select Case Board(Rank2, File2)
Case 0: PieceName2 = "Empty"
Case bRook: PieceName2 = "Rook"
Case bKnight: PieceName2 = "Knight"
Case bBishop: PieceName2 = "Bishop"
Case bQueen: PieceName2 = "Queen"
Case bKing: PieceName2 = "King"
Case bPawn: PieceName2 = "Pawn"
Case wRook: PieceName2 = "Rook"
Case wKnight: PieceName2 = "Knight"
Case wBishop: PieceName2 = "Bishop"
Case wQueen: PieceName2 = "Queen"
Case wKing: PieceName2 = "King"
Case wPawn: PieceName2 = "Pawn"
End Select

Log = Log + XCoord + ":" + Trim$(Str$(File1)) + " to " + XCoord2 + ":" + Trim$(Str$(File2)) + vbNewLine
End Sub

Private Sub Timer1_Timer()
Timer1.Enabled = False

Select Case Mode
Case "OfflinePC": If Turn = 0 Then Call PerformAI
Case "OfflineCC": PerformAI
End Select

Timer1.Enabled = True
End Sub
