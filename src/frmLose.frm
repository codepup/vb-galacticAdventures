VERSION 5.00
Begin VB.Form frmLose 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Game Over!"
   ClientHeight    =   3030
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   4665
   Icon            =   "frmLose.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "frmLose.frx":1272
   ScaleHeight     =   3030
   ScaleWidth      =   4665
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdMenu 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Continue"
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
      Left            =   1320
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   2160
      Width           =   2055
   End
   Begin VB.Label lblLose 
      Alignment       =   2  'Center
      BackColor       =   &H00000000&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   1815
      Left            =   720
      TabIndex        =   1
      Top             =   120
      Width           =   3135
   End
End
Attribute VB_Name = "frmLose"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'=======================================================================================================================================================================
'Matthew Lal
'June 5, 2010
'Lose Screen
'Displays the game over screen.
'=============================================================================================================================
Option Explicit

Private Sub cmdMenu_Click()
    'Resets fuel amount for new game.
    frmGalaxy1.intFuel = 0
    Unload frmGalaxy1

    'Closes the screen and returns to the home.
    Unload frmLose
    frmHome.Show
End Sub

Private Sub Form_Load()
    'Encouraging comment for the user to play again.
    lblLose.Caption = "Don't give up now Captain " & frmGalaxy1.strPlayerName & "!"
End Sub
