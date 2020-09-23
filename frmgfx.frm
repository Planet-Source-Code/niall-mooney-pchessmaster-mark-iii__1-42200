VERSION 5.00
Begin VB.Form frmgfx 
   Caption         =   "Form for Graphics storage"
   ClientHeight    =   4455
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   5445
   LinkTopic       =   "Form1"
   ScaleHeight     =   4455
   ScaleWidth      =   5445
   StartUpPosition =   3  'Windows Default
   Begin VB.PictureBox Picture1 
      AutoSize        =   -1  'True
      Height          =   1980
      Left            =   120
      Picture         =   "frmgfx.frx":0000
      ScaleHeight     =   128
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   384
      TabIndex        =   0
      Top             =   2280
      Width           =   5820
   End
   Begin VB.Image Image1 
      Height          =   7800
      Left            =   3960
      Picture         =   "frmgfx.frx":2A07
      Stretch         =   -1  'True
      Top             =   120
      Width           =   10260
   End
End
Attribute VB_Name = "frmgfx"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Load()
Me.Visible = False
End Sub
