VERSION 5.00
Begin VB.Form frmMarsQuiz 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Galactic Adventures - Refueling On Mars"
   ClientHeight    =   8415
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   10200
   Icon            =   "frmMarsQuiz.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "frmMarsQuiz.frx":1272
   ScaleHeight     =   8415
   ScaleWidth      =   10200
   StartUpPosition =   2  'CenterScreen
   Begin VB.OptionButton optDummyMars 
      Caption         =   "Dummy"
      Height          =   855
      Left            =   8640
      TabIndex        =   29
      Top             =   6480
      Value           =   -1  'True
      Visible         =   0   'False
      Width           =   1575
   End
   Begin VB.Frame fraMarsQ4 
      BackColor       =   &H80000012&
      Caption         =   "Question 4"
      ForeColor       =   &H00FFFFFF&
      Height          =   3615
      Left            =   3360
      TabIndex        =   18
      Top             =   3840
      Width           =   3255
      Begin VB.OptionButton optMarsQ4A1 
         BackColor       =   &H80000012&
         Caption         =   "Carbonite Rocks"
         ForeColor       =   &H00FFFFFF&
         Height          =   495
         Left            =   240
         TabIndex        =   27
         Top             =   840
         Width           =   1935
      End
      Begin VB.OptionButton optMarsQ4A3 
         BackColor       =   &H80000012&
         Caption         =   "Plate Tectonics"
         ForeColor       =   &H00FFFFFF&
         Height          =   495
         Left            =   240
         TabIndex        =   26
         Top             =   2280
         Width           =   1935
      End
      Begin VB.OptionButton optMarsQ4A2 
         BackColor       =   &H80000012&
         Caption         =   "Kelvin-Helmholtz Mechanism"
         ForeColor       =   &H00FFFFFF&
         Height          =   495
         Left            =   240
         TabIndex        =   25
         Top             =   1560
         Width           =   1935
      End
      Begin VB.OptionButton optMarsQ4A4 
         BackColor       =   &H80000012&
         Caption         =   "Greenhouse Effect"
         ForeColor       =   &H00FFFFFF&
         Height          =   495
         Left            =   240
         TabIndex        =   24
         Top             =   3000
         Width           =   1935
      End
      Begin VB.Label lblMarsQuizQuestion4 
         BackColor       =   &H80000012&
         Caption         =   "How does CO2 heat up a planet?"
         ForeColor       =   &H00FFFFFF&
         Height          =   495
         Left            =   120
         TabIndex        =   28
         Top             =   240
         Width           =   2775
      End
   End
   Begin VB.Frame fraMarsQ2 
      BackColor       =   &H80000012&
      Caption         =   "Question 2"
      ForeColor       =   &H00FFFFFF&
      Height          =   3615
      Left            =   120
      TabIndex        =   17
      Top             =   3840
      Width           =   3135
      Begin VB.OptionButton optMarsQ2A1 
         BackColor       =   &H80000012&
         Caption         =   "Atmosphere"
         ForeColor       =   &H00FFFFFF&
         Height          =   495
         Left            =   360
         TabIndex        =   22
         Top             =   840
         Width           =   1935
      End
      Begin VB.OptionButton optMarsQ2A2 
         BackColor       =   &H80000012&
         Caption         =   "Rocks"
         ForeColor       =   &H00FFFFFF&
         Height          =   495
         Left            =   360
         TabIndex        =   21
         Top             =   1560
         Width           =   1935
      End
      Begin VB.OptionButton optMarsQ2A3 
         BackColor       =   &H80000012&
         Caption         =   "Aliens"
         ForeColor       =   &H00FFFFFF&
         Height          =   495
         Left            =   360
         TabIndex        =   20
         Top             =   2280
         Width           =   1935
      End
      Begin VB.OptionButton optMarsQ2A4 
         BackColor       =   &H80000012&
         Caption         =   "Ice"
         ForeColor       =   &H00FFFFFF&
         Height          =   495
         Left            =   360
         TabIndex        =   19
         Top             =   3000
         Width           =   1935
      End
      Begin VB.Label lblMarsQuizQuestion3 
         BackColor       =   &H80000012&
         Caption         =   "What of Mars leaked into space causing Mars to ""die""?"
         ForeColor       =   &H00FFFFFF&
         Height          =   495
         Left            =   120
         TabIndex        =   23
         Top             =   240
         Width           =   2895
      End
   End
   Begin VB.Frame fraMarsQ5 
      BackColor       =   &H80000012&
      Caption         =   "Question 5"
      ForeColor       =   &H00FFFFFF&
      Height          =   3615
      Left            =   6720
      TabIndex        =   13
      Top             =   240
      Width           =   3255
      Begin VB.OptionButton optMarsQ5A1 
         BackColor       =   &H80000012&
         Caption         =   "True"
         ForeColor       =   &H00FFFFFF&
         Height          =   495
         Left            =   240
         TabIndex        =   15
         Top             =   840
         Width           =   1935
      End
      Begin VB.OptionButton optMarsQ5A2 
         BackColor       =   &H80000012&
         Caption         =   "False"
         ForeColor       =   &H00FFFFFF&
         Height          =   495
         Left            =   240
         TabIndex        =   14
         Top             =   1560
         Width           =   1935
      End
      Begin VB.Label lblMarsQuizQuestion5 
         BackColor       =   &H80000012&
         Caption         =   "Mars could of sustained life:"
         ForeColor       =   &H00FFFFFF&
         Height          =   495
         Left            =   120
         TabIndex        =   16
         Top             =   240
         Width           =   2655
      End
   End
   Begin VB.CommandButton cmdMarsQuizRefuel 
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
      Height          =   615
      Left            =   8160
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   7560
      Width           =   1815
   End
   Begin VB.OptionButton optMarsQ3A3 
      BackColor       =   &H80000012&
      Caption         =   "Plate Tectonics"
      ForeColor       =   &H00FFFFFF&
      Height          =   495
      Left            =   3600
      TabIndex        =   3
      Top             =   2520
      Width           =   1935
   End
   Begin VB.OptionButton optMarsQ3A2 
      BackColor       =   &H80000012&
      Caption         =   "Life"
      ForeColor       =   &H00FFFFFF&
      Height          =   495
      Left            =   3600
      TabIndex        =   2
      Top             =   1800
      Width           =   1935
   End
   Begin VB.OptionButton optMarsQ3A4 
      BackColor       =   &H80000012&
      Caption         =   "Carbon Dioxide"
      ForeColor       =   &H00FFFFFF&
      Height          =   495
      Left            =   3600
      TabIndex        =   1
      Top             =   3240
      Width           =   1935
   End
   Begin VB.OptionButton optMarsQ3A1 
      BackColor       =   &H80000012&
      Caption         =   "Oxygen"
      ForeColor       =   &H00FFFFFF&
      Height          =   495
      Left            =   3600
      TabIndex        =   0
      Top             =   1080
      Width           =   1935
   End
   Begin VB.Frame fraMarsQ1 
      BackColor       =   &H80000012&
      Caption         =   "Question1"
      ForeColor       =   &H00FFFFFF&
      Height          =   3615
      Left            =   120
      TabIndex        =   5
      Top             =   240
      Width           =   3135
      Begin VB.OptionButton optMarsQ1A1 
         BackColor       =   &H00000000&
         Caption         =   "Venus"
         ForeColor       =   &H00FFFFFF&
         Height          =   495
         Left            =   360
         TabIndex        =   12
         Top             =   840
         Width           =   1935
      End
      Begin VB.OptionButton optMarsQ1A2 
         BackColor       =   &H80000012&
         Caption         =   "Mercury"
         ForeColor       =   &H00FFFFFF&
         Height          =   495
         Left            =   360
         TabIndex        =   11
         Top             =   1560
         Width           =   1935
      End
      Begin VB.OptionButton optMarsQ1A3 
         BackColor       =   &H80000012&
         Caption         =   "Pluto"
         ForeColor       =   &H00FFFFFF&
         Height          =   495
         Left            =   360
         TabIndex        =   10
         Top             =   3000
         Width           =   1935
      End
      Begin VB.OptionButton optMarsQ1A4 
         BackColor       =   &H80000012&
         Caption         =   "Earth"
         ForeColor       =   &H00FFFFFF&
         Height          =   495
         Left            =   360
         TabIndex        =   9
         Top             =   2280
         Width           =   1935
      End
      Begin VB.Label lblMarsQuizQuestion1 
         BackColor       =   &H00000000&
         Caption         =   "Mars can be considered the older brother of what planet?"
         ForeColor       =   &H00FFFFFF&
         Height          =   495
         Index           =   0
         Left            =   120
         TabIndex        =   6
         Top             =   240
         Width           =   2895
      End
   End
   Begin VB.Frame fraMarsQ3 
      BackColor       =   &H80000012&
      Caption         =   "Question 3"
      ForeColor       =   &H00FFFFFF&
      Height          =   3615
      Left            =   3360
      TabIndex        =   7
      Top             =   240
      Width           =   3255
      Begin VB.Label lblMarsQuizQuestion2 
         BackColor       =   &H80000012&
         Caption         =   "What did Mars lack that the Earth had? *Controlled CO2 levels*"
         ForeColor       =   &H00FFFFFF&
         Height          =   495
         Left            =   120
         TabIndex        =   8
         Top             =   240
         Width           =   3015
      End
   End
End
Attribute VB_Name = "frmMarsQuiz"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'=======================================================================================================================================================================
'Matthew Lal
'June 5, 2010
'Mars Quiz
'The user refuels to continue the adventure by answer the questions correctly.
'=============================================================================================================================
Option Explicit

Private Sub cmdMarsQuizRefuel_Click()
   'Determines what answers are correct and what answers are not.
    If optMarsQ1A4.Value = True Then
    frmGalaxy3.intScoreQ1 = frmGalaxy3.intScoreQ1 + 1
        Else
        frmGalaxy3.intScoreQ1 = frmGalaxy3.intScoreQ1 - 1
    End If
    
    If optMarsQ2A1.Value = True Then
    frmGalaxy3.intScoreQ2 = frmGalaxy3.intScoreQ2 + 1
        Else
        frmGalaxy3.intScoreQ2 = frmGalaxy3.intScoreQ2 - 1
    End If
    
    If optMarsQ3A3.Value = True Then
    frmGalaxy3.intScoreQ3 = frmGalaxy3.intScoreQ3 + 1
        Else
        frmGalaxy3.intScoreQ3 = frmGalaxy3.intScoreQ3 - 1
    End If
    
    If optMarsQ4A4.Value = True Then
    frmGalaxy3.intScoreQ4 = frmGalaxy3.intScoreQ4 + 1
        Else
        frmGalaxy3.intScoreQ4 = frmGalaxy3.intScoreQ4 - 1
    End If
    
    If optMarsQ5A1.Value = True Then
    frmGalaxy3.intScoreQ5 = frmGalaxy3.intScoreQ5 + 1
        Else
        frmGalaxy3.intScoreQ5 = frmGalaxy3.intScoreQ5 - 1
    End If
    
    'Determines the users overall score for the current quiz.
    frmGalaxy3.intScore = frmGalaxy3.intScoreQ1 + frmGalaxy3.intScoreQ2 + frmGalaxy3.intScoreQ3 + frmGalaxy3.intScoreQ4 + frmGalaxy3.intScoreQ5
    
    'Adds the users score to the fuel.
    frmGalaxy1.intFuel = frmGalaxy1.intFuel + frmGalaxy3.intScore
        
    'Closes the quiz and returns to the galaxy map.
    Unload frmMarsQuiz
    frmGalaxy3.Show
    
    'Displays the users score.
    frmGalaxy3.intScore = MsgBox("10 gallons of fuel used at landing, and " & frmGalaxy3.intScore & " gallons of fuel retrieved.")
    
End Sub

Private Sub Form_Load()
    'Resets the variables to avoid logic errors.
    frmGalaxy3.intScore = 0
    frmGalaxy3.intScoreQ1 = 0
    frmGalaxy3.intScoreQ2 = 0
    frmGalaxy3.intScoreQ3 = 0
    frmGalaxy3.intScoreQ4 = 0
    frmGalaxy3.intScoreQ5 = 0
End Sub

