VERSION 5.00
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "MSWINSCK.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "RICHTX32.OCX"
Begin VB.Form frmChess3Server 
   BackColor       =   &H00FF8080&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Ni-Star Enterprises - ChessMASTER Mark III: Server Mode - Niall Mooney"
   ClientHeight    =   7800
   ClientLeft      =   150
   ClientTop       =   840
   ClientWidth     =   10260
   Icon            =   "Chess3Serve.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   520
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   684
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command3 
      Caption         =   "Go"
      Height          =   255
      Left            =   9600
      TabIndex        =   75
      Top             =   6840
      Width           =   615
   End
   Begin VB.TextBox txtUserName 
      Height          =   285
      Left            =   7920
      TabIndex        =   74
      Top             =   6840
      Width           =   1575
   End
   Begin VB.ComboBox cmbBoard 
      Height          =   315
      Left            =   4080
      Style           =   2  'Dropdown List
      TabIndex        =   72
      Top             =   6960
      Width           =   2895
   End
   Begin VB.ComboBox cmbPieces 
      Height          =   315
      Left            =   4080
      Locked          =   -1  'True
      Style           =   2  'Dropdown List
      TabIndex        =   71
      Top             =   6600
      Width           =   2895
   End
   Begin VB.ComboBox cmbBacky 
      Height          =   315
      Left            =   4080
      Style           =   2  'Dropdown List
      TabIndex        =   70
      Top             =   7320
      Width           =   2895
   End
   Begin VB.Timer Timer1 
      Left            =   1920
      Top             =   480
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Exe"
      Height          =   255
      Left            =   2610
      TabIndex        =   45
      Top             =   4725
      Width           =   495
   End
   Begin VB.ComboBox cmbTool 
      Height          =   315
      Left            =   90
      Style           =   2  'Dropdown List
      TabIndex        =   44
      Top             =   4695
      Width           =   2415
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Say"
      Height          =   255
      Left            =   3450
      TabIndex        =   43
      Top             =   7485
      Width           =   495
   End
   Begin VB.TextBox TxtSend 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00FF8080&
      BeginProperty Font 
         Name            =   "Small Fonts"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   90
      TabIndex        =   42
      Top             =   7485
      Width           =   3255
   End
   Begin RichTextLib.RichTextBox RTMain 
      Height          =   2295
      Left            =   90
      TabIndex        =   41
      Top             =   5085
      Width           =   3855
      _ExtentX        =   6800
      _ExtentY        =   4048
      _Version        =   393217
      BackColor       =   16744576
      Enabled         =   -1  'True
      ReadOnly        =   -1  'True
      ScrollBars      =   2
      Appearance      =   0
      AutoVerbMenu    =   -1  'True
      TextRTF         =   $"Chess3Serve.frx":08CA
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Small Fonts"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.PictureBox picColour2 
      Appearance      =   0  'Flat
      BackColor       =   &H00FF0000&
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   3570
      ScaleHeight     =   15
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   23
      TabIndex        =   40
      Top             =   4725
      Width           =   375
   End
   Begin VB.PictureBox picColour1 
      Appearance      =   0  'Flat
      BackColor       =   &H000000FF&
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   3210
      ScaleHeight     =   15
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   23
      TabIndex        =   39
      Top             =   4725
      Width           =   375
   End
   Begin VB.PictureBox picCanvas 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H00FF8080&
      ForeColor       =   &H80000008&
      Height          =   3030
      Left            =   105
      ScaleHeight     =   200
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   255
      TabIndex        =   38
      Top             =   1590
      Width           =   3855
      Begin VB.PictureBox picIcon 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         ForeColor       =   &H80000008&
         Height          =   510
         Index           =   11
         Left            =   360
         Picture         =   "Chess3Serve.frx":094A
         ScaleHeight     =   32
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   32
         TabIndex        =   69
         Top             =   1440
         Visible         =   0   'False
         Width           =   510
      End
      Begin VB.PictureBox picMask 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         ForeColor       =   &H80000008&
         Height          =   510
         Index           =   11
         Left            =   960
         Picture         =   "Chess3Serve.frx":118C
         ScaleHeight     =   32
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   32
         TabIndex        =   68
         Top             =   1440
         Visible         =   0   'False
         Width           =   510
      End
      Begin VB.PictureBox picIcon 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         ForeColor       =   &H80000008&
         Height          =   510
         Index           =   10
         Left            =   360
         Picture         =   "Chess3Serve.frx":1256
         ScaleHeight     =   32
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   32
         TabIndex        =   67
         Top             =   1320
         Visible         =   0   'False
         Width           =   510
      End
      Begin VB.PictureBox picMask 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         ForeColor       =   &H80000008&
         Height          =   510
         Index           =   10
         Left            =   960
         Picture         =   "Chess3Serve.frx":1A98
         ScaleHeight     =   32
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   32
         TabIndex        =   66
         Top             =   1320
         Visible         =   0   'False
         Width           =   510
      End
      Begin VB.PictureBox picIcon 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         ForeColor       =   &H80000008&
         Height          =   510
         Index           =   9
         Left            =   360
         Picture         =   "Chess3Serve.frx":1B62
         ScaleHeight     =   32
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   32
         TabIndex        =   65
         Top             =   1200
         Visible         =   0   'False
         Width           =   510
      End
      Begin VB.PictureBox picMask 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         ForeColor       =   &H80000008&
         Height          =   510
         Index           =   9
         Left            =   960
         Picture         =   "Chess3Serve.frx":1DE4
         ScaleHeight     =   32
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   32
         TabIndex        =   64
         Top             =   1200
         Visible         =   0   'False
         Width           =   510
      End
      Begin VB.PictureBox picMask 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         ForeColor       =   &H80000008&
         Height          =   510
         Index           =   8
         Left            =   960
         Picture         =   "Chess3Serve.frx":1EAE
         ScaleHeight     =   32
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   32
         TabIndex        =   63
         Top             =   1080
         Visible         =   0   'False
         Width           =   510
      End
      Begin VB.PictureBox picIcon 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         ForeColor       =   &H80000008&
         Height          =   510
         Index           =   8
         Left            =   360
         Picture         =   "Chess3Serve.frx":1F78
         ScaleHeight     =   32
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   32
         TabIndex        =   62
         Top             =   1080
         Visible         =   0   'False
         Width           =   510
      End
      Begin VB.PictureBox picMask 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         ForeColor       =   &H80000008&
         Height          =   510
         Index           =   7
         Left            =   960
         Picture         =   "Chess3Serve.frx":27BA
         ScaleHeight     =   32
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   32
         TabIndex        =   61
         Top             =   960
         Visible         =   0   'False
         Width           =   510
      End
      Begin VB.PictureBox picIcon 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         ForeColor       =   &H80000008&
         Height          =   510
         Index           =   7
         Left            =   360
         Picture         =   "Chess3Serve.frx":2884
         ScaleHeight     =   32
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   32
         TabIndex        =   60
         Top             =   960
         Visible         =   0   'False
         Width           =   510
      End
      Begin VB.PictureBox picIcon 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         ForeColor       =   &H80000008&
         Height          =   510
         Index           =   6
         Left            =   360
         Picture         =   "Chess3Serve.frx":30C6
         ScaleHeight     =   32
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   32
         TabIndex        =   59
         Top             =   840
         Visible         =   0   'False
         Width           =   510
      End
      Begin VB.PictureBox picIcon 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         ForeColor       =   &H80000008&
         Height          =   510
         Index           =   5
         Left            =   360
         Picture         =   "Chess3Serve.frx":3348
         ScaleHeight     =   32
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   32
         TabIndex        =   57
         Top             =   720
         Visible         =   0   'False
         Width           =   510
      End
      Begin VB.PictureBox picIcon 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         ForeColor       =   &H80000008&
         Height          =   510
         Index           =   4
         Left            =   360
         Picture         =   "Chess3Serve.frx":35CA
         ScaleHeight     =   32
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   32
         TabIndex        =   55
         Top             =   600
         Visible         =   0   'False
         Width           =   510
      End
      Begin VB.PictureBox picIcon 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         ForeColor       =   &H80000008&
         Height          =   510
         Index           =   3
         Left            =   360
         Picture         =   "Chess3Serve.frx":384C
         ScaleHeight     =   32
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   32
         TabIndex        =   53
         Top             =   480
         Visible         =   0   'False
         Width           =   510
      End
      Begin VB.PictureBox picMask 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         ForeColor       =   &H80000008&
         Height          =   510
         Index           =   6
         Left            =   960
         Picture         =   "Chess3Serve.frx":3ACE
         ScaleHeight     =   32
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   32
         TabIndex        =   58
         Top             =   840
         Visible         =   0   'False
         Width           =   510
      End
      Begin VB.PictureBox picMask 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         ForeColor       =   &H80000008&
         Height          =   510
         Index           =   5
         Left            =   960
         Picture         =   "Chess3Serve.frx":3B98
         ScaleHeight     =   32
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   32
         TabIndex        =   56
         Top             =   720
         Visible         =   0   'False
         Width           =   510
      End
      Begin VB.PictureBox picMask 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         ForeColor       =   &H80000008&
         Height          =   510
         Index           =   4
         Left            =   960
         Picture         =   "Chess3Serve.frx":3C62
         ScaleHeight     =   32
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   32
         TabIndex        =   54
         Top             =   600
         Visible         =   0   'False
         Width           =   510
      End
      Begin VB.PictureBox picMask 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         ForeColor       =   &H80000008&
         Height          =   510
         Index           =   3
         Left            =   960
         Picture         =   "Chess3Serve.frx":3D2C
         ScaleHeight     =   32
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   32
         TabIndex        =   52
         Top             =   480
         Visible         =   0   'False
         Width           =   510
      End
      Begin VB.PictureBox picIcon 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         ForeColor       =   &H80000008&
         Height          =   510
         Index           =   2
         Left            =   360
         Picture         =   "Chess3Serve.frx":3DF6
         ScaleHeight     =   32
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   32
         TabIndex        =   51
         Top             =   360
         Visible         =   0   'False
         Width           =   510
      End
      Begin VB.PictureBox picMask 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         ForeColor       =   &H80000008&
         Height          =   510
         Index           =   2
         Left            =   960
         Picture         =   "Chess3Serve.frx":4078
         ScaleHeight     =   32
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   32
         TabIndex        =   50
         Top             =   360
         Visible         =   0   'False
         Width           =   510
      End
      Begin VB.PictureBox picMask 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         ForeColor       =   &H80000008&
         Height          =   510
         Index           =   1
         Left            =   960
         Picture         =   "Chess3Serve.frx":4142
         ScaleHeight     =   32
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   32
         TabIndex        =   49
         Top             =   240
         Visible         =   0   'False
         Width           =   510
      End
      Begin VB.PictureBox picIcon 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         ForeColor       =   &H80000008&
         Height          =   510
         Index           =   1
         Left            =   360
         Picture         =   "Chess3Serve.frx":420C
         ScaleHeight     =   32
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   32
         TabIndex        =   48
         Top             =   240
         Visible         =   0   'False
         Width           =   510
      End
      Begin VB.PictureBox picMask 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         ForeColor       =   &H80000008&
         Height          =   510
         Index           =   0
         Left            =   960
         Picture         =   "Chess3Serve.frx":448E
         ScaleHeight     =   32
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   32
         TabIndex        =   47
         Top             =   120
         Visible         =   0   'False
         Width           =   510
      End
      Begin VB.PictureBox picIcon 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         ForeColor       =   &H80000008&
         Height          =   510
         Index           =   0
         Left            =   360
         Picture         =   "Chess3Serve.frx":4710
         ScaleHeight     =   32
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   32
         TabIndex        =   46
         Top             =   120
         Visible         =   0   'False
         Width           =   510
      End
      Begin VB.Shape Shape1 
         BorderStyle     =   3  'Dot
         Height          =   975
         Left            =   2040
         Top             =   1200
         Visible         =   0   'False
         Width           =   735
      End
      Begin VB.Line Line1 
         BorderStyle     =   3  'Dot
         Visible         =   0   'False
         X1              =   128
         X2              =   72
         Y1              =   64
         Y2              =   128
      End
   End
   Begin MSWinsockLib.Winsock WsServe 
      Left            =   1440
      Top             =   480
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
      RemotePort      =   49974
      LocalPort       =   4997
   End
   Begin VB.PictureBox picPiece 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      Height          =   780
      Index           =   1
      Left            =   1335
      Picture         =   "Chess3Serve.frx":4992
      ScaleHeight     =   48
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   48
      TabIndex        =   33
      Top             =   1605
      Visible         =   0   'False
      Width           =   780
   End
   Begin VB.PictureBox picPiece 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      Height          =   780
      Index           =   2
      Left            =   1695
      Picture         =   "Chess3Serve.frx":64D4
      ScaleHeight     =   48
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   48
      TabIndex        =   32
      Top             =   1605
      Visible         =   0   'False
      Width           =   780
   End
   Begin VB.PictureBox picPiece 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      Height          =   780
      Index           =   3
      Left            =   2055
      Picture         =   "Chess3Serve.frx":8016
      ScaleHeight     =   48
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   48
      TabIndex        =   31
      Top             =   1605
      Visible         =   0   'False
      Width           =   780
   End
   Begin VB.PictureBox picPiece 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      Height          =   780
      Index           =   4
      Left            =   2415
      Picture         =   "Chess3Serve.frx":9B58
      ScaleHeight     =   48
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   48
      TabIndex        =   30
      Top             =   1605
      Visible         =   0   'False
      Width           =   780
   End
   Begin VB.PictureBox picPiece 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      Height          =   780
      Index           =   5
      Left            =   2775
      Picture         =   "Chess3Serve.frx":B69A
      ScaleHeight     =   48
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   48
      TabIndex        =   29
      Top             =   1605
      Visible         =   0   'False
      Width           =   780
   End
   Begin VB.PictureBox picPiece 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      Height          =   780
      Index           =   6
      Left            =   3135
      Picture         =   "Chess3Serve.frx":D1DC
      ScaleHeight     =   48
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   48
      TabIndex        =   28
      Top             =   1605
      Visible         =   0   'False
      Width           =   780
   End
   Begin VB.PictureBox picPiece 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      Height          =   780
      Index           =   11
      Left            =   1335
      Picture         =   "Chess3Serve.frx":ED1E
      ScaleHeight     =   48
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   48
      TabIndex        =   27
      Top             =   1845
      Visible         =   0   'False
      Width           =   780
   End
   Begin VB.PictureBox picPiece 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      Height          =   780
      Index           =   12
      Left            =   1695
      Picture         =   "Chess3Serve.frx":10860
      ScaleHeight     =   48
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   48
      TabIndex        =   26
      Top             =   1845
      Visible         =   0   'False
      Width           =   780
   End
   Begin VB.PictureBox picPiece 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      Height          =   780
      Index           =   13
      Left            =   2055
      Picture         =   "Chess3Serve.frx":123A2
      ScaleHeight     =   48
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   48
      TabIndex        =   25
      Top             =   1845
      Visible         =   0   'False
      Width           =   780
   End
   Begin VB.PictureBox picPiece 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      Height          =   780
      Index           =   14
      Left            =   2415
      Picture         =   "Chess3Serve.frx":13EE4
      ScaleHeight     =   48
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   48
      TabIndex        =   24
      Top             =   1845
      Visible         =   0   'False
      Width           =   780
   End
   Begin VB.PictureBox picPiece 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      Height          =   780
      Index           =   15
      Left            =   2775
      Picture         =   "Chess3Serve.frx":15A26
      ScaleHeight     =   48
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   48
      TabIndex        =   23
      Top             =   1845
      Visible         =   0   'False
      Width           =   780
   End
   Begin VB.PictureBox picPiece 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      Height          =   780
      Index           =   16
      Left            =   3135
      Picture         =   "Chess3Serve.frx":17568
      ScaleHeight     =   48
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   48
      TabIndex        =   22
      Top             =   1845
      Visible         =   0   'False
      Width           =   780
   End
   Begin VB.PictureBox picPiece 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      Height          =   780
      Index           =   21
      Left            =   1335
      Picture         =   "Chess3Serve.frx":190AA
      ScaleHeight     =   48
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   48
      TabIndex        =   21
      Top             =   2085
      Visible         =   0   'False
      Width           =   780
   End
   Begin VB.PictureBox picPiece 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      Height          =   780
      Index           =   22
      Left            =   1695
      Picture         =   "Chess3Serve.frx":19274
      ScaleHeight     =   48
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   48
      TabIndex        =   20
      Top             =   2085
      Visible         =   0   'False
      Width           =   780
   End
   Begin VB.PictureBox picPiece 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      Height          =   780
      Index           =   23
      Left            =   2055
      Picture         =   "Chess3Serve.frx":1943E
      ScaleHeight     =   48
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   48
      TabIndex        =   19
      Top             =   2085
      Visible         =   0   'False
      Width           =   780
   End
   Begin VB.PictureBox picPiece 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      Height          =   780
      Index           =   24
      Left            =   2415
      Picture         =   "Chess3Serve.frx":19608
      ScaleHeight     =   48
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   48
      TabIndex        =   18
      Top             =   2085
      Visible         =   0   'False
      Width           =   780
   End
   Begin VB.PictureBox picPiece 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      Height          =   780
      Index           =   25
      Left            =   2775
      Picture         =   "Chess3Serve.frx":197D2
      ScaleHeight     =   48
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   48
      TabIndex        =   17
      Top             =   2085
      Visible         =   0   'False
      Width           =   780
   End
   Begin VB.PictureBox picPiece 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      Height          =   780
      Index           =   26
      Left            =   3135
      Picture         =   "Chess3Serve.frx":1999C
      ScaleHeight     =   48
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   48
      TabIndex        =   16
      Top             =   2085
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
   Begin VB.PictureBox picEffector 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00FF0000&
      Height          =   780
      Index           =   1
      Left            =   8880
      Picture         =   "Chess3Serve.frx":19B66
      ScaleHeight     =   48
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   48
      TabIndex        =   15
      Top             =   2880
      Visible         =   0   'False
      Width           =   780
   End
   Begin VB.PictureBox picEffector 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00FF0000&
      Height          =   780
      Index           =   0
      Left            =   8880
      Picture         =   "Chess3Serve.frx":1A8A8
      ScaleHeight     =   48
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   48
      TabIndex        =   14
      Top             =   2640
      Visible         =   0   'False
      Width           =   780
   End
   Begin VB.PictureBox picFrontBuffer 
      AutoRedraw      =   -1  'True
      Height          =   5820
      Left            =   4320
      ScaleHeight     =   384
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   384
      TabIndex        =   0
      Top             =   480
      Width           =   5820
   End
   Begin VB.Label Label5 
      BackStyle       =   0  'Transparent
      Caption         =   "I/O"
      Height          =   255
      Left            =   480
      TabIndex        =   83
      Top             =   1110
      Width           =   255
   End
   Begin VB.Shape Shape2 
      BackColor       =   &H000000FF&
      Height          =   255
      Left            =   240
      Top             =   1080
      Width           =   735
   End
   Begin VB.Label RemoteUN 
      BackStyle       =   0  'Transparent
      Caption         =   "..."
      Height          =   255
      Left            =   7920
      TabIndex        =   82
      Top             =   7440
      Width           =   2295
   End
   Begin VB.Label Label4 
      BackStyle       =   0  'Transparent
      Caption         =   "UserName:"
      Height          =   255
      Left            =   7080
      TabIndex        =   81
      Top             =   7440
      Width           =   855
   End
   Begin VB.Label lblClientIP 
      BackStyle       =   0  'Transparent
      Caption         =   "..."
      Height          =   255
      Left            =   7920
      TabIndex        =   80
      Top             =   7200
      Width           =   2295
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "Client IP:"
      Height          =   255
      Left            =   7080
      TabIndex        =   79
      Top             =   7200
      Width           =   735
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "UserName: "
      Height          =   255
      Left            =   7080
      TabIndex        =   78
      Top             =   6840
      Width           =   855
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Server IP:"
      Height          =   255
      Left            =   7080
      TabIndex        =   77
      Top             =   6600
      Width           =   735
   End
   Begin VB.Label lblIP 
      BackStyle       =   0  'Transparent
      Caption         =   "LOCAL INTERNET PROTOCOL"
      Height          =   255
      Left            =   7920
      TabIndex        =   76
      Top             =   6600
      Width           =   2415
   End
   Begin VB.Label Label6 
      BackStyle       =   0  'Transparent
      Caption         =   "Chess Graphics Set"
      Height          =   255
      Left            =   4080
      TabIndex        =   73
      Top             =   6390
      Width           =   3255
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
      Left            =   7560
      TabIndex        =   37
      Top             =   120
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
      Left            =   8280
      TabIndex        =   36
      Top             =   120
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
      Left            =   9000
      TabIndex        =   35
      Top             =   120
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
      Left            =   9720
      TabIndex        =   34
      Top             =   120
      Width           =   255
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
      Left            =   6795
      TabIndex        =   13
      Top             =   120
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
      Left            =   6075
      TabIndex        =   12
      Top             =   120
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
      Left            =   5355
      TabIndex        =   11
      Top             =   120
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
      Left            =   4635
      TabIndex        =   10
      Top             =   120
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
      Left            =   4080
      TabIndex        =   9
      Top             =   5760
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
      Left            =   4080
      TabIndex        =   8
      Top             =   5040
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
      Left            =   4080
      TabIndex        =   7
      Top             =   4320
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
      Left            =   4080
      TabIndex        =   6
      Top             =   3600
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
      Left            =   4080
      TabIndex        =   5
      Top             =   2880
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
      Left            =   4080
      TabIndex        =   4
      Top             =   2160
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
      Left            =   4080
      TabIndex        =   3
      Top             =   1440
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
      Left            =   4080
      TabIndex        =   2
      Top             =   720
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
      Height          =   1335
      Left            =   120
      TabIndex        =   1
      Top             =   120
      Width           =   3855
   End
   Begin VB.Image Image1 
      Height          =   7800
      Left            =   0
      Picture         =   "Chess3Serve.frx":1B523
      Stretch         =   -1  'True
      Top             =   0
      Width           =   10260
   End
   Begin VB.Menu mnuFile 
      Caption         =   "&File"
      Begin VB.Menu mnuExit 
         Caption         =   "Exit"
      End
   End
   Begin VB.Menu mnuServe 
      Caption         =   "Service"
      Begin VB.Menu mnuRARS 
         Caption         =   "Restart Game"
      End
      Begin VB.Menu mnuARsef 
         Caption         =   "Disconnect"
      End
      Begin VB.Menu mnuRules 
         Caption         =   "Rules"
         Checked         =   -1  'True
      End
   End
   Begin VB.Menu mnuGraphics 
      Caption         =   "Graphics"
      Begin VB.Menu mnuDangerShow 
         Caption         =   "Show Units in Danger"
         Checked         =   -1  'True
      End
   End
   Begin VB.Menu mnuAbout 
      Caption         =   "About"
   End
End
Attribute VB_Name = "frmChess3Server"
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

Dim Lx As Integer, Ly As Integer, SelX As Integer, SelY As Integer, Selected As Boolean
Dim Board(7, 7) As Byte
Dim BoardAI(7, 7) As Byte
Dim Turn As Integer
Dim AIType As Integer
Dim MenuChoiceRate As Single
Dim PriorityMenu As Integer
Dim Mode As String 'OfflinePP, OfflinePC, OfflineCC, Client, Server
Dim DrawTool As Integer
Dim StartX As Integer, StartY As Integer, EndX As Integer, EndY As Integer
Dim Username As String
Dim LastSend As Long
Dim LastMsgPART As String

Private Declare Function BitBlt Lib "gdi32" (ByVal hDestDC As Long, ByVal X As Long, ByVal Y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal XSrc As Long, ByVal YSrc As Long, ByVal dwRop As Long) As Long
Private Declare Function GetTickCount Lib "kernel32" () As Long
Private Declare Function AlphaBlending Lib "msimg32.dll" Alias "AlphaBlend" (ByVal hdcDest As Long, ByVal nXOriginDest As Long, ByVal nYOriginDest As Long, ByVal nWidthDest As Long, ByVal nHeightDest As Long, ByVal hdcSrc As Long, ByVal nXOriginSrc As Long, ByVal nYOriginSrc As Long, ByVal nWidthSrc As Long, ByVal nHeightSrc As Long, ByVal BF As Long) As Long
Private Declare Function DrawTransparent Lib "msimg32.dll" Alias "TransparentBlt" (ByVal hdcDest As Long, ByVal nXOriginDest As Long, ByVal nYOriginDest As Long, ByVal nWidthDest As Long, ByVal nHeightDest As Long, ByVal hdcSrc As Long, ByVal nXOriginSrc As Long, ByVal nYOriginSrc As Long, ByVal nWidthSrc As Long, ByVal nHeightSrc As Long, ByVal crTransparent As Long) As Long
Private Declare Function Ellipse Lib "gdi32" (ByVal hdc As Long, ByVal X1 As Long, ByVal Y1 As Long, ByVal X2 As Long, ByVal Y2 As Long) As Long

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
If Mode <> "Server" Then Exit Sub
WsServe.SendData "InstantMsg" + Chr$(27) + Trim$(txtUserName.Text) + " says: " + TxtSend.Text + vbNewLine
RTMain.Text = Trim$(txtUserName.Text) + " says: " + TxtSend.Text + vbNewLine + RTMain.Text
Call MarkNetworkActivity
End Sub

Private Sub Command2_Click()
If cmbTool.ListIndex = 0 Then
picCanvas.Cls
If Mode = "Server" Then WsServe.SendData "PicMsg" + Chr$(27) + "Clear"
Call MarkNetworkActivity
End If
End Sub

Private Sub Command3_Click()
If Mode = "Server" Then InitiateNewGame
End Sub

Private Sub Form_Initialize()
On Error Resume Next

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

cmbTool.AddItem "Clear Canvas"
cmbTool.AddItem "Line"
cmbTool.AddItem "Box Filled"
cmbTool.AddItem "Box Unfilled"
cmbTool.AddItem "Ellipse Filled"
cmbTool.AddItem "Ellipse Unfilled"
cmbTool.AddItem "Lightning"
cmbTool.AddItem "Banana"
cmbTool.AddItem "Extatic Smiley"
cmbTool.AddItem "Pleased Smiley"
cmbTool.AddItem "It 's Okay Smiley"
cmbTool.AddItem "Displeased Smiley"
cmbTool.AddItem "Crying Smiley"
cmbTool.AddItem "Logo 1"
cmbTool.AddItem "Logo 2"
cmbTool.AddItem "N'yer"
cmbTool.AddItem "UK flag"
cmbTool.AddItem "US flag"
cmbTool.ListIndex = 1

cmbBacky.AddItem "Default Sky Blue Background"
cmbBacky.ListIndex = 0

lblIP.Caption = WsServe.LocalIP
txtUserName.Text = WsServe.LocalHostName
WsServe.Listen

If Err Then MsgBox "There was an error initiating ChessMASTER Mark III, please try again if this fails then reboot your computer.", vbCritical + vbOKOnly: End
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
If MsgBox("Are you certain that you wish to leave at this time", vbQuestion + vbYesNo, "NSE Provisional Chess 3") = vbYes Then
Unload Me
End
End If
Cancel = 1
End Sub

Public Sub InitiateNewGame()
Dim Indexia As Integer, Jedexia As Integer

WsServe.SendData "InitiateNewGame"
Call MarkNetworkActivity

For Indexia = 0 To 7
For Jedexia = 0 To 7
Board(Jedexia, Indexia) = 0
Next Jedexia
Next Indexia

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
End Sub

Public Sub MakeMove(Rank1 As Integer, File1 As Integer, Rank2 As Integer, File2 As Integer)
If Mode <> "Server" Then Exit Sub

If mnuRules.Checked = True Then
If MoveAlright(Rank1, File1, Rank2, File2) Then
Board(Rank2, File2) = Board(Rank1, File1)
Board(Rank1, File1) = 0
If Turn = 0 Then Turn = 1 Else Turn = 0
Selected = False
RenderBoard
End If
Else
Board(Rank2, File2) = Board(Rank1, File1)
Board(Rank1, File1) = 0
If Turn = 0 Then Turn = 1 Else Turn = 0
Selected = False
RenderBoard
End If

WsServe.SendData "MakeMove" + Chr$(27) + Trim$(Str$(Rank1)) + Chr$(27) + Trim$(Str$(File1)) + Chr$(27) + Trim$(Str$(Rank2)) + Chr$(27) + Trim$(Str$(File2))
Call MarkNetworkActivity

Select Case Mode

End Select

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

If Colour = "B" And Mode = "Server" Then
MoveAlright = False
Exit Function
End If

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
Dim X&, Y&

If Board(Sx, Sy) < 10 Then Colour$ = "B"
If Board(Sx, Sy) > 10 Then Colour$ = "W"
If Board(tx, ty) < 10 And Board(tx, ty) > 0 Then ColourT = "B"
If Board(tx, ty) > 10 Then ColourT = "W"

If Colour = "B" And Mode = "Server" Then
MoveLegal = False
Exit Function
End If

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

If Err Then MsgBox Err.Description, vbCritical, "Serious Error"
End Function

Private Sub mnuAbout_Click()
frmAbout.Show
End Sub

Private Sub mnuARsef_Click()
WsServe.Close
End Sub

Private Sub mnuExit_Click()
Unload Me
End Sub

Private Sub mnuRARS_Click()
If Mode <> "Server" Then Exit Sub
WsServe.SendData "InitiateNewGame"
Call MarkNetworkActivity
End Sub

Private Sub mnuRules_Click()
If Mode <> "Server" Then Exit Sub

If mnuRules.Checked Then mnuRules.Checked = False Else mnuRules.Checked = True
If mnuRules.Checked = True Then WsServe.SendData "RulesOFF": Call MarkNetworkActivity
End Sub

Private Sub picCanvas_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)

If cmbTool.ListIndex = 1 Then 'Line Drawing
StartX = X
StartY = Y
Line1.X1 = StartX
Line1.Y1 = StartY
Line1.X2 = X
Line1.Y2 = Y
Line1.Visible = True
End If

If cmbTool.ListIndex = 2 Or cmbTool.ListIndex = 3 Then  'Box drawing
StartX = X
StartY = Y
Shape1.Left = StartX
Shape1.Top = StartY
Shape1.Width = 0
Shape1.Height = 0
Shape1.Visible = True
Shape1.Shape = 0
End If

If cmbTool.ListIndex = 4 Or cmbTool.ListIndex = 5 Then  'Ellipse drawing
StartX = X
StartY = Y
Shape1.Left = StartX
Shape1.Top = StartY
Shape1.Width = 0
Shape1.Height = 0
Shape1.Visible = True
Shape1.Shape = 2
End If

End Sub

Private Sub picCanvas_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

If Y > 200 Or Y < 0 Or X > 255 Or X < 0 Then Exit Sub

If cmbTool.ListIndex = 1 Then 'Line Drawing
Line1.X1 = StartX
Line1.Y1 = StartY
Line1.X2 = X
Line1.Y2 = Y
End If

If cmbTool.ListIndex = 2 Or cmbTool.ListIndex = 3 Then  'Box drawing
If StartX < X Then Shape1.Left = StartX Else Shape1.Left = X
If StartY < Y Then Shape1.Top = StartY Else Shape1.Top = Y
Shape1.Width = Abs(StartX - X)
Shape1.Height = Abs(StartY - Y)
End If

If cmbTool.ListIndex = 4 Or cmbTool.ListIndex = 5 Then  'Ellipse drawing
If StartX < X Then Shape1.Left = StartX Else Shape1.Left = X
If StartY < Y Then Shape1.Top = StartY Else Shape1.Top = Y
Shape1.Width = Abs(StartX - X)
Shape1.Height = Abs(StartY - Y)
End If

End Sub

Private Sub picCanvas_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)

If X > 255 Then X = 255
If Y > 200 Then Y = 200
If X < 0 Then X = 0
If Y < 0 Then Y = 0

If cmbTool.ListIndex = 1 Then 'Line Drawing
EndX = X
EndY = Y
Line1.Visible = False
If Button = vbLeftButton Then picCanvas.ForeColor = picColour1.BackColor Else picCanvas.ForeColor = picColour2.BackColor
If Mode = "Server" And SendingAllowed Then WsServe.SendData "PicMsg" + Chr$(27) + "Line" + Chr$(27) + Trim$(Str$(StartX)) + Chr$(27) + Trim$(Str$(StartY)) + Chr$(27) + Trim$(Str$(EndX)) + Chr$(27) + Trim$(Str$(EndY)) + Chr$(27) + Trim$(Str$(picCanvas.ForeColor))
picCanvas.Line (StartX, StartY)-(EndX, EndY)
End If

If cmbTool.ListIndex = 2 Then   'Box drawing, Filled
EndX = X
EndY = Y
Shape1.Visible = False
If Button = vbLeftButton Then picCanvas.ForeColor = picColour1.BackColor Else picCanvas.ForeColor = picColour2.BackColor
If Button = vbLeftButton Then picCanvas.FillColor = picColour1.BackColor Else picCanvas.FillColor = picColour2.BackColor
If Mode = "Server" And SendingAllowed Then WsServe.SendData "PicMsg" + Chr$(27) + "BF" + Chr$(27) + Trim$(Str$(StartX)) + Chr$(27) + Trim$(Str$(StartY)) + Chr$(27) + Trim$(Str$(EndX)) + Chr$(27) + Trim$(Str$(EndY)) + Chr$(27) + Trim$(Str$(picCanvas.ForeColor))
picCanvas.Line (StartX, StartY)-(EndX, EndY), , BF
End If
If cmbTool.ListIndex = 3 Then   'Box drawing, Unfilled
EndX = X
EndY = Y
Shape1.Visible = False
If Button = vbLeftButton Then picCanvas.ForeColor = picColour1.BackColor Else picCanvas.ForeColor = picColour2.BackColor
If Mode = "Server" And SendingAllowed Then WsServe.SendData "PicMsg" + Chr$(27) + "BUF" + Chr$(27) + Trim$(Str$(StartX)) + Chr$(27) + Trim$(Str$(StartY)) + Chr$(27) + Trim$(Str$(EndX)) + Chr$(27) + Trim$(Str$(EndY)) + Chr$(27) + Trim$(Str$(picCanvas.ForeColor))
picCanvas.Line (StartX, StartY)-(EndX, EndY), , B
End If

If cmbTool.ListIndex = 4 Then    'Ellipse drawing,  Filled
EndX = X
EndY = Y
Shape1.Visible = False
If Button = vbLeftButton Then picCanvas.ForeColor = picColour1.BackColor Else picCanvas.ForeColor = picColour2.BackColor
If Button = vbLeftButton Then picCanvas.FillColor = picColour1.BackColor Else picCanvas.FillColor = picColour2.BackColor
picCanvas.FillStyle = 0
If Mode = "Server" And SendingAllowed Then WsServe.SendData "PicMsg" + Chr$(27) + "EF" + Chr$(27) + Trim$(Str$(StartX)) + Chr$(27) + Trim$(Str$(StartY)) + Chr$(27) + Trim$(Str$(EndX)) + Chr$(27) + Trim$(Str$(EndY)) + Chr$(27) + Trim$(Str$(picCanvas.ForeColor))
Ellipse picCanvas.hdc, StartX, StartY, EndX, EndY
End If
If cmbTool.ListIndex = 5 Then    'Ellipse drawing, Unfilled
EndX = X
EndY = Y
Shape1.Visible = False
If Button = vbLeftButton Then picCanvas.ForeColor = picColour1.BackColor Else picCanvas.ForeColor = picColour2.BackColor
picCanvas.FillStyle = 1
If Mode = "Server" And SendingAllowed Then WsServe.SendData "PicMsg" + Chr$(27) + "EUF" + Chr$(27) + Trim$(Str$(StartX)) + Chr$(27) + Trim$(Str$(StartY)) + Chr$(27) + Trim$(Str$(EndX)) + Chr$(27) + Trim$(Str$(EndY)) + Chr$(27) + Trim$(Str$(picCanvas.ForeColor))
Ellipse picCanvas.hdc, StartX, StartY, EndX, EndY
End If

If cmbTool.ListIndex = 6 Then 'Lightning
EndX = X - 16
EndY = Y - 16
BitBlt picCanvas.hdc, EndX, EndY, 32, 32, picMask(0).hdc, 0, 0, vbMergePaint
BitBlt picCanvas.hdc, EndX, EndY, 32, 32, picIcon(0).hdc, 0, 0, vbSrcAnd 'vbMergePaint
If Mode = "Server" And SendingAllowed Then WsServe.SendData "PicMsg" + Chr$(27) + "Light" + Chr$(27) + Trim$(Str$(EndX)) + Chr$(27) + Trim$(Str$(EndY))
End If
If cmbTool.ListIndex = 7 Then 'Banana
EndX = X - 16
EndY = Y - 16
BitBlt picCanvas.hdc, EndX, EndY, 32, 32, picMask(1).hdc, 0, 0, vbMergePaint
BitBlt picCanvas.hdc, EndX, EndY, 32, 32, picIcon(1).hdc, 0, 0, vbSrcAnd 'vbMergePaint
If Mode = "Server" And SendingAllowed Then WsServe.SendData "PicMsg" + Chr$(27) + "Banan" + Chr$(27) + Trim$(Str$(EndX)) + Chr$(27) + Trim$(Str$(EndY))
End If
If cmbTool.ListIndex = 8 Then 'Extatic Smiley
EndX = X - 16
EndY = Y - 16
BitBlt picCanvas.hdc, EndX, EndY, 32, 32, picMask(2).hdc, 0, 0, vbMergePaint
BitBlt picCanvas.hdc, EndX, EndY, 32, 32, picIcon(2).hdc, 0, 0, vbSrcAnd 'vbMergePaint
If Mode = "Server" And SendingAllowed Then WsServe.SendData "PicMsg" + Chr$(27) + "Extat" + Chr$(27) + Trim$(Str$(EndX)) + Chr$(27) + Trim$(Str$(EndY))
End If
If cmbTool.ListIndex = 9 Then 'Pleased Smiley
EndX = X - 16
EndY = Y - 16
BitBlt picCanvas.hdc, EndX, EndY, 32, 32, picMask(3).hdc, 0, 0, vbMergePaint
BitBlt picCanvas.hdc, EndX, EndY, 32, 32, picIcon(3).hdc, 0, 0, vbSrcAnd 'vbMergePaint
If Mode = "Server" And SendingAllowed Then WsServe.SendData "PicMsg" + Chr$(27) + "Pleas" + Chr$(27) + Trim$(Str$(EndX)) + Chr$(27) + Trim$(Str$(EndY))
End If
If cmbTool.ListIndex = 10 Then 'C'est Okay Smiley
EndX = X - 16
EndY = Y - 16
BitBlt picCanvas.hdc, EndX, EndY, 32, 32, picMask(4).hdc, 0, 0, vbMergePaint
BitBlt picCanvas.hdc, EndX, EndY, 32, 32, picIcon(4).hdc, 0, 0, vbSrcAnd 'vbMergePaint
If Mode = "Server" And SendingAllowed Then WsServe.SendData "PicMsg" + Chr$(27) + "Okay" + Chr$(27) + Trim$(Str$(EndX)) + Chr$(27) + Trim$(Str$(EndY))
End If
If cmbTool.ListIndex = 11 Then 'Displeased Smiley
EndX = X - 16
EndY = Y - 16
BitBlt picCanvas.hdc, EndX, EndY, 32, 32, picMask(5).hdc, 0, 0, vbMergePaint
BitBlt picCanvas.hdc, EndX, EndY, 32, 32, picIcon(5).hdc, 0, 0, vbSrcAnd 'vbMergePaint
If Mode = "Server" And SendingAllowed Then WsServe.SendData "PicMsg" + Chr$(27) + "UHappy" + Chr$(27) + Trim$(Str$(EndX)) + Chr$(27) + Trim$(Str$(EndY))
End If
If cmbTool.ListIndex = 12 Then 'Displeased Smiley
EndX = X - 16
EndY = Y - 16
BitBlt picCanvas.hdc, EndX, EndY, 32, 32, picMask(6).hdc, 0, 0, vbMergePaint
BitBlt picCanvas.hdc, EndX, EndY, 32, 32, picIcon(6).hdc, 0, 0, vbSrcAnd 'vbMergePaint
If Mode = "Server" And SendingAllowed Then WsServe.SendData "PicMsg" + Chr$(27) + "Crying" + Chr$(27) + Trim$(Str$(EndX)) + Chr$(27) + Trim$(Str$(EndY))
End If
If cmbTool.ListIndex = 13 Then 'Crying Smiley
EndX = X - 16
EndY = Y - 16
BitBlt picCanvas.hdc, EndX, EndY, 32, 32, picMask(7).hdc, 0, 0, vbMergePaint
BitBlt picCanvas.hdc, EndX, EndY, 32, 32, picIcon(7).hdc, 0, 0, vbSrcAnd 'vbMergePaint
If Mode = "Server" And SendingAllowed Then WsServe.SendData "PicMsg" + Chr$(27) + "Logo1" + Chr$(27) + Trim$(Str$(EndX)) + Chr$(27) + Trim$(Str$(EndY))
End If
If cmbTool.ListIndex = 14 Then 'Logo 1
EndX = X - 16
EndY = Y - 16
BitBlt picCanvas.hdc, EndX, EndY, 32, 32, picMask(8).hdc, 0, 0, vbMergePaint
BitBlt picCanvas.hdc, EndX, EndY, 32, 32, picIcon(8).hdc, 0, 0, vbSrcAnd 'vbMergePaint
If Mode = "Server" And SendingAllowed Then WsServe.SendData "PicMsg" + Chr$(27) + "Logo2" + Chr$(27) + Trim$(Str$(EndX)) + Chr$(27) + Trim$(Str$(EndY))
End If
If cmbTool.ListIndex = 15 Then 'Logo 2
EndX = X - 16
EndY = Y - 16
BitBlt picCanvas.hdc, EndX, EndY, 32, 32, picMask(9).hdc, 0, 0, vbMergePaint
BitBlt picCanvas.hdc, EndX, EndY, 32, 32, picIcon(9).hdc, 0, 0, vbSrcAnd 'vbMergePaint
If Mode = "Server" And SendingAllowed Then WsServe.SendData "PicMsg" + Chr$(27) + "Nyer" + Chr$(27) + Trim$(Str$(EndX)) + Chr$(27) + Trim$(Str$(EndY))
End If
If cmbTool.ListIndex = 16 Then 'UK Flag
EndX = X - 16
EndY = Y - 16
BitBlt picCanvas.hdc, EndX, EndY, 32, 32, picMask(10).hdc, 0, 0, vbMergePaint
BitBlt picCanvas.hdc, EndX, EndY, 32, 32, picIcon(10).hdc, 0, 0, vbSrcAnd 'vbMergePaint
If Mode = "Server" And SendingAllowed Then WsServe.SendData "PicMsg" + Chr$(27) + "UK" + Chr$(27) + Trim$(Str$(EndX)) + Chr$(27) + Trim$(Str$(EndY))
End If
If cmbTool.ListIndex = 17 Then 'US Flag
EndX = X - 16
EndY = Y - 16
BitBlt picCanvas.hdc, EndX, EndY, 32, 32, picMask(11).hdc, 0, 0, vbMergePaint
BitBlt picCanvas.hdc, EndX, EndY, 32, 32, picIcon(11).hdc, 0, 0, vbSrcAnd 'vbMergePaint
If Mode = "Server" And SendingAllowed Then WsServe.SendData "PicMsg" + Chr$(27) + "US" + Chr$(27) + Trim$(Str$(EndX)) + Chr$(27) + Trim$(Str$(EndY))
End If

If Mode = "Server" Then Call MarkNetworkActivity

picCanvas.Refresh
End Sub

Private Sub picColour1_Click()
CD.ShowColor
picColour1.BackColor = CD.Color
End Sub

Private Sub picColour2_Click()
CD.ShowColor
picColour2.BackColor = CD.Color
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
Case "Client"
Status = Status + "Online Client Mode" + vbNewLine
Case "Server"
Status = Status + "Online Server Mode" + vbNewLine
Case ""
Status = Status + "ChessMASTER Mark III" + vbNewLine
Status = Status + "Waiting for a Client to Connect" + vbNewLine
End Select

Select Case Turn
Case 0
Status = Status + "Waiting for " + RemoteUN.Caption
Case 1
Status = Status + "Your Turn, " + txtUserName.Text
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
Status = Status + vbNewLine + PieceName + " at " + XCoord + ":" + Str$(SelY + 1) + " To " + XCoord2 + ":" + Str$(Ly + 1) + " (" + PieceName2 + ")"
picFrontBuffer.ToolTipText = PieceName + " at " + XCoord + ":" + Str$(SelY + 1) + " To " + XCoord2 + ":" + Str$(Ly + 1) + " (" + PieceName2 + ")"
If MoveAlright(SelX, SelY, Lx, Ly) Then Status = Status + vbNewLine + "~Move Legal~": picFrontBuffer.ToolTipText = "[Legal] " + picFrontBuffer.ToolTipText
If mnuRules.Checked = False Then Status = Status + vbNewLine + "~No Move Law~": picFrontBuffer.ToolTipText = "[Legal] " + picFrontBuffer.ToolTipText
Else
Status = Status + vbNewLine + PieceName2 + " at " + XCoord2 + ", " + Str(Ly + 1)
End If

lblStatus.Caption = Status$
End Sub

Private Sub EndGame()
Dim X&, Y&

Mode = ""
Turn = -1

For X = 0 To 7
For Y = 0 To 7
Board(X, Y) = 0
BoardAI(X, Y) = 0
Next Y
Next X

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

Private Sub CheckGame()
Dim bKingX As Integer, bKingY As Integer
Dim wKingX As Integer, wKingY As Integer


End Sub

Private Sub Timer1_Timer()
Timer1.Interval = 0
Shape2.BackStyle = 0
End Sub

Private Sub txtUserName_Change()
Username = txtUserName.Text
If Mode = "Server" Then WsServe.SendData "ID" + Chr$(27) + Username: Call MarkNetworkActivity
End Sub

Private Sub WsServe_Close()
RTMain.Text = "-Networking Critical Error. The Client is Disconnected" + vbNewLine + RTMain.Text
MsgBox "Networking Critical Error. The Client has been disconnected or has choosen to 'disconnect in disgrace'. You are the Champion!", vbExclamation + vbOKOnly
Mode = ""
End Sub

Private Sub WsServe_ConnectionRequest(ByVal requestID As Long)
RTMain.Text = "-Connected to " + WsServe.RemoteHostIP + "." + vbNewLine + RTMain.Text
Mode = "Server"
WsServe.Close
WsServe.Accept requestID
WsServe.SendData "ID" + Chr$(27) + Username
Beep
WsServe.SendData "InstantMsg" + Chr$(27) + vbNewLine + "Ni-Star Enterprises" + vbNewLine + "ChessMASTER Mark III" + vbNewLine + "Server Mode build 3056" + vbNewLine + "Welcome To ChessMASTER" + vbNewLine + "Mark III Online"
Call MarkNetworkActivity
RTMain.Text = "-Press GO to initiate the GameState..." + vbNewLine + RTMain.Text
RTMain.Text = "----------------" + vbNewLine + "-Welcome to Ni-Star Enterprises ChessMASTER Mark III" + vbNewLine + "----------------" + vbNewLine + RTMain.Text
End Sub

Private Sub WsServe_DataArrival(ByVal bytesTotal As Long)
Dim ReceivedData As String
Dim Segment() As String

WsServe.GetData ReceivedData

Segment() = Split(ReceivedData, Chr$(27))

If Segment(0) = "Game" Then
Select Case Segment(1)
Case "StartNew"
InitiateNewGame
End Select
End If

If Segment(0) = "InstantMsg" Then
RTMain.Text = Segment(1) + RTMain.Text
End If

If Segment(0) = "PicMsg" Then
Select Case Segment(1)
Case "Clear"
picCanvas.Cls
Case "Line"
picCanvas.Line (Val(Segment(2)), Val(Segment(3)))-(Val(Segment(4)), Val(Segment(5))), Segment(6)
Case "BF"
picCanvas.Line (Val(Segment(2)), Val(Segment(3)))-(Val(Segment(4)), Val(Segment(5))), Segment(6), BF
Case "BUF"
picCanvas.Line (Val(Segment(2)), Val(Segment(3)))-(Val(Segment(4)), Val(Segment(5))), Segment(6), B
Case "EF"
picCanvas.FillStyle = 0
picCanvas.ForeColor = Val(Segment(6))
picCanvas.FillColor = Val(Segment(6))
Ellipse picCanvas.hdc, Val(Segment(2)), Val(Segment(3)), Val(Segment(4)), Val(Segment(5))
Case "EUF"
picCanvas.FillStyle = 1
picCanvas.ForeColor = Segment(6)
Ellipse picCanvas.hdc, Val(Segment(2)), Val(Segment(3)), Val(Segment(4)), Val(Segment(5))
Case "Light"
BitBlt picCanvas.hdc, Val(Segment(2)), Val(Segment(3)), 32, 32, picMask(0).hdc, 0, 0, vbMergePaint
BitBlt picCanvas.hdc, Val(Segment(2)), Val(Segment(3)), 32, 32, picIcon(0).hdc, 0, 0, vbSrcAnd 'vbMergePaint
Case "Banan"
BitBlt picCanvas.hdc, Val(Segment(2)), Val(Segment(3)), 32, 32, picMask(1).hdc, 0, 0, vbMergePaint
BitBlt picCanvas.hdc, Val(Segment(2)), Val(Segment(3)), 32, 32, picIcon(1).hdc, 0, 0, vbSrcAnd 'vbMergePaint
Case "Extat"
BitBlt picCanvas.hdc, Val(Segment(2)), Val(Segment(3)), 32, 32, picMask(2).hdc, 0, 0, vbMergePaint
BitBlt picCanvas.hdc, Val(Segment(2)), Val(Segment(3)), 32, 32, picIcon(2).hdc, 0, 0, vbSrcAnd 'vbMergePaint
Case "Pleas"
BitBlt picCanvas.hdc, Val(Segment(2)), Val(Segment(3)), 32, 32, picMask(3).hdc, 0, 0, vbMergePaint
BitBlt picCanvas.hdc, Val(Segment(2)), Val(Segment(3)), 32, 32, picIcon(3).hdc, 0, 0, vbSrcAnd 'vbMergePaint
Case "Okay"
BitBlt picCanvas.hdc, Val(Segment(2)), Val(Segment(3)), 32, 32, picMask(4).hdc, 0, 0, vbMergePaint
BitBlt picCanvas.hdc, Val(Segment(2)), Val(Segment(3)), 32, 32, picIcon(4).hdc, 0, 0, vbSrcAnd 'vbMergePaint
Case "UHappy"
BitBlt picCanvas.hdc, Val(Segment(2)), Val(Segment(3)), 32, 32, picMask(5).hdc, 0, 0, vbMergePaint
BitBlt picCanvas.hdc, Val(Segment(2)), Val(Segment(3)), 32, 32, picIcon(5).hdc, 0, 0, vbSrcAnd 'vbMergePaint
Case "Crying"
BitBlt picCanvas.hdc, Val(Segment(2)), Val(Segment(3)), 32, 32, picMask(6).hdc, 0, 0, vbMergePaint
BitBlt picCanvas.hdc, Val(Segment(2)), Val(Segment(3)), 32, 32, picIcon(6).hdc, 0, 0, vbSrcAnd 'vbMergePaint
Case "Logo1"
BitBlt picCanvas.hdc, Val(Segment(2)), Val(Segment(3)), 32, 32, picMask(7).hdc, 0, 0, vbMergePaint
BitBlt picCanvas.hdc, Val(Segment(2)), Val(Segment(3)), 32, 32, picIcon(7).hdc, 0, 0, vbSrcAnd 'vbMergePaint
Case "Logo2"
BitBlt picCanvas.hdc, Val(Segment(2)), Val(Segment(3)), 32, 32, picMask(8).hdc, 0, 0, vbMergePaint
BitBlt picCanvas.hdc, Val(Segment(2)), Val(Segment(3)), 32, 32, picIcon(8).hdc, 0, 0, vbSrcAnd 'vbMergePaint
Case "Nyer"
BitBlt picCanvas.hdc, Val(Segment(2)), Val(Segment(3)), 32, 32, picMask(9).hdc, 0, 0, vbMergePaint
BitBlt picCanvas.hdc, Val(Segment(2)), Val(Segment(3)), 32, 32, picIcon(9).hdc, 0, 0, vbSrcAnd 'vbMergePaint
Case "UK"
BitBlt picCanvas.hdc, Val(Segment(2)), Val(Segment(3)), 32, 32, picMask(10).hdc, 0, 0, vbMergePaint
BitBlt picCanvas.hdc, Val(Segment(2)), Val(Segment(3)), 32, 32, picIcon(10).hdc, 0, 0, vbSrcAnd 'vbMergePaint
Case "US"
BitBlt picCanvas.hdc, Val(Segment(2)), Val(Segment(3)), 32, 32, picMask(11).hdc, 0, 0, vbMergePaint
BitBlt picCanvas.hdc, Val(Segment(2)), Val(Segment(3)), 32, 32, picIcon(11).hdc, 0, 0, vbSrcAnd 'vbMergePaint
End Select
picCanvas.Refresh
End If

If Segment(0) = "ID" Then
RemoteUN.Caption = Segment(1)
lblClientIP.Caption = WsServe.RemoteHostIP
End If

If Segment(0) = "MakeMove" Then
Board(Val(Segment(3)), Val(Segment(4))) = Board(Val(Segment(1)), Val(Segment(2)))
Board(Val(Segment(1)), Val(Segment(2))) = 0
If Turn = 0 Then Turn = 1 Else Turn = 0
Selected = False
RenderBoard
End If

Call MarkNetworkActivity
End Sub

Private Sub WsServe_Error(ByVal Number As Integer, Description As String, ByVal Scode As Long, ByVal Source As String, ByVal HelpFile As String, ByVal HelpContext As Long, CancelDisplay As Boolean)
If WsServe.State <> sckConnected Then RTMain.Text = "Networking Critical Error." + vbNewLine + RTMain.Text Else RTMain.Text = "Networking Error" + vbNewLine + RTMain.Text
End Sub

Private Sub MarkNetworkActivity()
Timer1.Interval = 100
Shape2.BackStyle = 1
End Sub

Private Function SendingAllowed() As Boolean
If LastSend + 500 < GetTickCount Then
SendingAllowed = True
LastSend = GetTickCount
End If
End Function

