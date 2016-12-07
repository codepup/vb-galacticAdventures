VERSION 5.00
Begin VB.Form frmSaturn 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Galactic Adventures"
   ClientHeight    =   8415
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   10230
   Icon            =   "frmSaturn.frx":0000
   LinkTopic       =   "Form4"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "frmSaturn.frx":1272
   ScaleHeight     =   8415
   ScaleWidth      =   10230
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdSaturnRefuel 
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
      TabIndex        =   1
      Top             =   5280
      Width           =   1455
   End
   Begin VB.CommandButton cmdSaturnBack 
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
      TabIndex        =   0
      Top             =   5280
      Width           =   1455
   End
   Begin VB.Label lblSaturn 
      Alignment       =   2  'Center
      BackColor       =   &H00000000&
      Caption         =   "Saturn"
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
   Begin VB.Image imgSaturn 
      Height          =   2715
      Left            =   360
      Picture         =   "frmSaturn.frx":6AEF
      Top             =   1320
      Width           =   4500
   End
   Begin VB.Label lblSaturnInfo 
      BackColor       =   &H00000000&
      Caption         =   $"frmSaturn.frx":2E785
      ForeColor       =   &H00FFFFFF&
      Height          =   6855
      Left            =   5280
      TabIndex        =   2
      Top             =   1320
      Width           =   4695
   End
End
Attribute VB_Name = "frmSaturn"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'=======================================================================================================================================================================
'Matthew Lal
'June 5, 2010
'Saturn Information
'Displays all the information on planet Saturn.
'=============================================================================================================================
Option Explicit

Private Sub cmdSaturnBack_Click()
    'Closes the info screen and returns to the galaxy map.
    frmSaturn.Hide
    frmGalaxy3.Show
End Sub

Private Sub cmdSaturnRefuel_Click()
    'Closes the info screen and starts the quiz.
    frmSaturn.Hide
    frmSaturnQuiz.Show
End Sub

