VERSION 5.00
Begin VB.Form frmPluto 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Galactic Adventures"
   ClientHeight    =   8415
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   10230
   Icon            =   "frmPluto.frx":0000
   LinkTopic       =   "Form7"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "frmPluto.frx":1272
   ScaleHeight     =   8415
   ScaleWidth      =   10230
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdPlutoRefuel 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Refuel"
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
      Left            =   3480
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   6960
      Width           =   1455
   End
   Begin VB.CommandButton cmdPlutoBack 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Back"
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
      Left            =   480
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   6960
      Width           =   1455
   End
   Begin VB.Label lblPluto 
      Alignment       =   2  'Center
      BackColor       =   &H00000000&
      Caption         =   "Pluto"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   24
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   495
      Left            =   1680
      TabIndex        =   3
      Top             =   360
      Width           =   1575
   End
   Begin VB.Image imgPluto 
      Height          =   4890
      Left            =   120
      Picture         =   "frmPluto.frx":6AEF
      Stretch         =   -1  'True
      Top             =   1200
      Width           =   4890
   End
   Begin VB.Label lblPlutoInfo 
      BackColor       =   &H00000000&
      Caption         =   $"frmPluto.frx":A329
      ForeColor       =   &H00FFFFFF&
      Height          =   6615
      Left            =   5280
      TabIndex        =   0
      Top             =   1200
      Width           =   4695
   End
End
Attribute VB_Name = "frmPluto"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'=======================================================================================================================================================================
'Matthew Lal
'June 5, 2010
'Pluto Information
'Displays all the information on planet Pluto.
'=============================================================================================================================
Option Explicit

Private Sub cmdPlutoBack_Click()
    'Closes the info screen and returns to the galaxy map.
    frmPluto.Hide
    frmGalaxy3.Show
End Sub

Private Sub cmdPlutoRefuel_Click()
    'Closes the info screen and starts the quiz.
    frmPluto.Hide
    frmPlutoQuiz.Show
End Sub
