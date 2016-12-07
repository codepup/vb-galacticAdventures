VERSION 5.00
Begin VB.Form frmEarth 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Galactic Adventures"
   ClientHeight    =   8415
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   10230
   Icon            =   "frmEarth.frx":0000
   LinkTopic       =   "Form12"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "frmEarth.frx":1272
   ScaleHeight     =   8415
   ScaleWidth      =   10230
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdEarthBack 
      Appearance      =   0  'Flat
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
      Left            =   360
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   6960
      Width           =   1455
   End
   Begin VB.CommandButton cmdEarthRefuel 
      Appearance      =   0  'Flat
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
      Left            =   3360
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   6960
      Width           =   1455
   End
   Begin VB.Label lblEarthInfo 
      BackColor       =   &H00000000&
      Caption         =   $"frmEarth.frx":6AEF
      ForeColor       =   &H00FFFFFF&
      Height          =   6615
      Left            =   5280
      TabIndex        =   3
      Top             =   1200
      Width           =   4695
   End
   Begin VB.Image imgEarth 
      Height          =   4890
      Left            =   120
      Picture         =   "frmEarth.frx":73E4
      Stretch         =   -1  'True
      Top             =   1200
      Width           =   4890
   End
   Begin VB.Label lblEarth 
      Alignment       =   2  'Center
      BackColor       =   &H00000000&
      Caption         =   "Earth"
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
      TabIndex        =   2
      Top             =   360
      Width           =   1575
   End
End
Attribute VB_Name = "frmEarth"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'=======================================================================================================================================================================
'Matthew Lal
'June 5, 2010
'Earth Information
'Displays all information on the planet Earth.
'=======================================================================================================================================================================
Option Explicit

Private Sub cmdEarthBack_Click()
    'Closes the Earth Info screen and displays the galaxy map.
    Unload frmEarth
    frmGalaxy3.Show
End Sub

Private Sub cmdEarthRefuel_Click()
    'Closes the Earth Info screen and starts the quiz.
    Unload frmEarth
    frmEarthQuiz.Show
End Sub

