VERSION 5.00
Begin VB.Form frmVenus 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Galactic Adventures"
   ClientHeight    =   8415
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   10230
   Icon            =   "frmVenus.frx":0000
   LinkTopic       =   "Form11"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "frmVenus.frx":1272
   ScaleHeight     =   8415
   ScaleWidth      =   10230
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdVenusRefuel 
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
      Top             =   6960
      Width           =   1455
   End
   Begin VB.CommandButton cmdVenusBack 
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
      TabIndex        =   0
      Top             =   6960
      Width           =   1455
   End
   Begin VB.Label lblVenus 
      Alignment       =   2  'Center
      BackColor       =   &H00000000&
      Caption         =   "Venus"
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
   Begin VB.Image imgVenus 
      Height          =   4890
      Left            =   120
      Picture         =   "frmVenus.frx":6AEF
      Stretch         =   -1  'True
      Top             =   1200
      Width           =   4890
   End
   Begin VB.Label lblVenusInfo 
      BackColor       =   &H00000000&
      Caption         =   $"frmVenus.frx":48A5D
      ForeColor       =   &H00FFFFFF&
      Height          =   6135
      Left            =   5280
      TabIndex        =   2
      Top             =   1200
      Width           =   4695
   End
End
Attribute VB_Name = "frmVenus"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'=======================================================================================================================================================================
'Matthew Lal
'June 5, 2010
'Venus Information
'Displays all the information on planet Venus.
'=============================================================================================================================
Option Explicit

Private Sub cmdVenusBack_Click()
    'Closes the info screen and returns to the galaxy map.
    frmVenus.Hide
    frmGalaxy3.Show
End Sub

Private Sub cmdVenusRefuel_Click()
    'Closes the info screen and starts the quiz.
    frmVenus.Hide
    frmVenusQuiz.Show
End Sub
