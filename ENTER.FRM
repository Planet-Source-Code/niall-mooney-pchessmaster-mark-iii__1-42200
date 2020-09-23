VERSION 5.00
Begin VB.Form frmEnter 
   BackColor       =   &H00000000&
   BorderStyle     =   0  'None
   Caption         =   "Ni-Star Enterprises - ChessMASTER Mark III Entré"
   ClientHeight    =   2490
   ClientLeft      =   2790
   ClientTop       =   3150
   ClientWidth     =   4095
   BeginProperty Font 
      Name            =   "Verdana"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   ForeColor       =   &H0000FFFF&
   Icon            =   "ENTER.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "ENTER.frx":030A
   ScaleHeight     =   166
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   273
   StartUpPosition =   2  'CenterScreen
End
Attribute VB_Name = "frmEnter"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim BoardPicNumber As Long
Dim Selected As Boolean

Private Declare Function ReleaseCapture Lib "user32" () As Long
Private Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long

Private Sub Form_Load()
BoardPicNumber = 1

Me.Show

Do
frmEnter.ZOrder 0
DoEvents
Loop Until Selected
End Sub

Private Sub Form_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
If Button = 1 Then
If X > 63 And X < 210 And Y > 47 And Y < 65 Then Unload Me: frmChess3Server.Show: Selected = True
If X > 63 And X < 210 And Y > 65 And Y < 88 Then Unload Me:  frmChess3Client.Show: Selected = True
'If X > 63 And X < 210 And Y > 88 And Y < 113 Then Unload Me:  frmSpectator.Show:selected=True
If X > 63 And X < 210 And Y > 113 And Y < 132 Then Unload Me:  frmChess3Offline.Show: Selected = True
End If
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
If X > 63 And X < 210 And Y > 47 And Y < 65 Then Me.Picture = LoadResPicture(114, 0)
If X > 63 And X < 210 And Y > 65 And Y < 88 Then Me.Picture = LoadResPicture(115, 0)
If X > 63 And X < 210 And Y > 88 And Y < 113 Then Me.Picture = LoadResPicture(116, 0)
If X > 63 And X < 210 And Y > 113 And Y < 132 Then Me.Picture = LoadResPicture(117, 0)

If Button = 1 Then
If X > 63 And X < 210 And Y > 47 And Y < 65 Then Unload Me: frmChess3Server.Show: Selected = True
If X > 63 And X < 210 And Y > 65 And Y < 88 Then Unload Me:  frmChess3Client.Show: Selected = True
'If X > 63 And X < 210 And Y > 88 And Y < 113 Then Unload Me:  frmSpectator.Show:selected=True
If X > 63 And X < 210 And Y > 113 And Y < 132 Then Unload Me:  frmChess3Offline.Show: Selected = True
End If

'Move Window
If Button = 1 Then ReleaseCapture: SendMessage Me.hwnd, &HA1, 2, 0&
End Sub
