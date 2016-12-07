VERSION 5.00
Begin VB.Form frmGalaxy2 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Galactic Adventures"
   ClientHeight    =   8415
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   10230
   Icon            =   "frmGalaxy2.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "frmGalaxy2.frx":1272
   ScaleHeight     =   8415
   ScaleWidth      =   10230
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdNext2 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Launch"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   7920
      MaskColor       =   &H00E0E0E0&
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   7800
      Width           =   1455
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Let's get a move on!"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   27.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1815
      Left            =   6480
      TabIndex        =   2
      Top             =   1320
      Width           =   2295
   End
   Begin VB.Image Image1 
      BorderStyle     =   1  'Fixed Single
      Height          =   6330
      Left            =   360
      Picture         =   "frmGalaxy2.frx":6AEF
      Stretch         =   -1  'True
      Top             =   120
      Width           =   9495
   End
   Begin VB.Label lblDialog2 
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Left            =   840
      TabIndex        =   1
      Top             =   6600
      Width           =   8535
   End
End
Attribute VB_Name = "frmGalaxy2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'=======================================================================================================================================================================
'Matthew Lal
'June 5, 2010
'Introduction Screen 2
'Puts the user name into a dialog screen, then continues into the game.
'=============================================================================================================================
Option Explicit

Private Sub cmdNext2_Click()
    'Closes the window and opens the galaxy map.
    frmGalaxy3.Show
    Unload frmGalaxy2
End Sub

Private Sub Form_Load()
    'Dosplays a message to the user with the users name.
    lblDialog2.Caption = "Let's get a move on! " & frmGalaxy1.strPlayerName & ", our mission is to gather intel about the missing crew. We'll lose fuel sooner or later so land and refuel at near by planets, and find the crew."
End Sub
