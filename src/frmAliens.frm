VERSION 5.00
Begin VB.Form frmAliens 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Galactic Adventures"
   ClientHeight    =   8445
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   10230
   Icon            =   "frmAliens.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "frmAliens.frx":1272
   ScaleHeight     =   8445
   ScaleWidth      =   10230
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton cmdLastChance 
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
      Left            =   4440
      MaskColor       =   &H00E0E0E0&
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   7800
      Width           =   1455
   End
   Begin VB.Label lblDialogAlien 
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   1  'Fixed Single
      Caption         =   $"frmAliens.frx":6AEF
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Left            =   480
      TabIndex        =   1
      Top             =   6600
      Width           =   9255
   End
   Begin VB.Image imgAlienPortrait 
      BorderStyle     =   1  'Fixed Single
      Height          =   6285
      Left            =   360
      Picture         =   "frmAliens.frx":6C1D
      Stretch         =   -1  'True
      Top             =   120
      Width           =   9495
   End
End
Attribute VB_Name = "frmAliens"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'=======================================================================================================================================================================
'Matthew Lal
'June 5, 2010
'Alien Dialog
'Last chance for the user to continue the game, and must answer one question.
'=======================================================================================================================================================================Option Explicit
Dim strLastChance As String

Private Sub cmdLastChance_Click()
    'Last question that determines whether the user wins or loses.
    strLastChance = InputBox("How did space come to be?")
    If strLastChance = "Big Bang" Then
        Unload frmAliens
        frmWin.Show
    ElseIf strLastChance = "big bang" Then
        Unload frmAliens
        frmWin.Show
    Else
        Unload frmAliens
        frmLose.Show
    End If
End Sub
