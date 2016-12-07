VERSION 5.00
Begin VB.Form frmVenusQuiz 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Galactic Adventures - Refueling On Venus"
   ClientHeight    =   8415
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   10200
   Icon            =   "frmVenusQuiz.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "frmVenusQuiz.frx":1272
   ScaleHeight     =   8415
   ScaleWidth      =   10200
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame fraVenusQ3 
      BackColor       =   &H80000012&
      Caption         =   "Question 3"
      ForeColor       =   &H00FFFFFF&
      Height          =   3615
      Left            =   3480
      TabIndex        =   23
      Top             =   240
      Width           =   3255
      Begin VB.OptionButton optVenusQ3A3 
         BackColor       =   &H80000012&
         Caption         =   "Earth"
         ForeColor       =   &H00FFFFFF&
         Height          =   495
         Left            =   360
         TabIndex        =   27
         Top             =   2280
         Width           =   1935
      End
      Begin VB.OptionButton optVenusQ3A2 
         BackColor       =   &H80000012&
         Caption         =   "Mercury"
         ForeColor       =   &H00FFFFFF&
         Height          =   495
         Left            =   360
         TabIndex        =   26
         Top             =   1560
         Width           =   1935
      End
      Begin VB.OptionButton optVenusQ3A4 
         BackColor       =   &H80000012&
         Caption         =   "Sun"
         ForeColor       =   &H00FFFFFF&
         Height          =   495
         Left            =   360
         TabIndex        =   25
         Top             =   3000
         Width           =   1935
      End
      Begin VB.OptionButton optVenusQ3A1 
         BackColor       =   &H80000012&
         Caption         =   "Mars"
         ForeColor       =   &H00FFFFFF&
         Height          =   495
         Left            =   360
         TabIndex        =   24
         Top             =   840
         Width           =   1935
      End
      Begin VB.Label lblVenusQuizQuestion2 
         BackColor       =   &H80000012&
         Caption         =   "Venus is simillar to what planet?"
         ForeColor       =   &H00FFFFFF&
         Height          =   495
         Left            =   120
         TabIndex        =   28
         Top             =   240
         Width           =   3015
      End
   End
   Begin VB.Frame fraVenusQ1 
      BackColor       =   &H80000012&
      Caption         =   "Question1"
      ForeColor       =   &H00FFFFFF&
      Height          =   3615
      Left            =   240
      TabIndex        =   17
      Top             =   240
      Width           =   3135
      Begin VB.OptionButton optVenusQ1A1 
         BackColor       =   &H00000000&
         Caption         =   "Venera 7"
         ForeColor       =   &H00FFFFFF&
         Height          =   495
         Left            =   360
         TabIndex        =   21
         Top             =   840
         Width           =   1935
      End
      Begin VB.OptionButton optVenusQ1A2 
         BackColor       =   &H80000012&
         Caption         =   "Apollo 4"
         ForeColor       =   &H00FFFFFF&
         Height          =   495
         Left            =   360
         TabIndex        =   20
         Top             =   1560
         Width           =   1935
      End
      Begin VB.OptionButton optVenusQ1A3 
         BackColor       =   &H80000012&
         Caption         =   "Mariner 2"
         ForeColor       =   &H00FFFFFF&
         Height          =   495
         Left            =   360
         TabIndex        =   19
         Top             =   2280
         Width           =   1935
      End
      Begin VB.OptionButton optVenusQ1A4 
         BackColor       =   &H80000012&
         Caption         =   "Voyager 2"
         ForeColor       =   &H00FFFFFF&
         Height          =   495
         Left            =   360
         TabIndex        =   18
         Top             =   3000
         Width           =   1935
      End
      Begin VB.Label lblVenusQuizQuestion1 
         BackColor       =   &H00000000&
         Caption         =   "The landing of what spacecraft solved the long awaited question of Venus?"
         ForeColor       =   &H00FFFFFF&
         Height          =   495
         Index           =   0
         Left            =   120
         TabIndex        =   22
         Top             =   240
         Width           =   2895
      End
   End
   Begin VB.Frame fraVenusQ5 
      BackColor       =   &H80000012&
      Caption         =   "Question 5"
      ForeColor       =   &H00FFFFFF&
      Height          =   3615
      Left            =   6840
      TabIndex        =   13
      Top             =   240
      Width           =   3255
      Begin VB.OptionButton optVenusQ5A1 
         BackColor       =   &H80000012&
         Caption         =   "True"
         ForeColor       =   &H00FFFFFF&
         Height          =   495
         Left            =   240
         TabIndex        =   15
         Top             =   840
         Width           =   1935
      End
      Begin VB.OptionButton optVenusQ5A2 
         BackColor       =   &H80000012&
         Caption         =   "False"
         ForeColor       =   &H00FFFFFF&
         Height          =   495
         Left            =   240
         TabIndex        =   14
         Top             =   1560
         Width           =   1935
      End
      Begin VB.Label lblVenusQuizQuestion5 
         BackColor       =   &H80000012&
         Caption         =   "Venus' days are longer than it's years."
         ForeColor       =   &H00FFFFFF&
         Height          =   495
         Left            =   120
         TabIndex        =   16
         Top             =   240
         Width           =   2655
      End
   End
   Begin VB.Frame fraVenusQ2 
      BackColor       =   &H80000012&
      Caption         =   "Question 2"
      ForeColor       =   &H00FFFFFF&
      Height          =   3615
      Left            =   240
      TabIndex        =   7
      Top             =   3840
      Width           =   3135
      Begin VB.OptionButton optVenusQ2A1 
         BackColor       =   &H80000012&
         Caption         =   "Atmosphere"
         ForeColor       =   &H00FFFFFF&
         Height          =   495
         Left            =   360
         TabIndex        =   11
         Top             =   840
         Width           =   1935
      End
      Begin VB.OptionButton optVenusQ2A2 
         BackColor       =   &H80000012&
         Caption         =   "Environment"
         ForeColor       =   &H00FFFFFF&
         Height          =   495
         Left            =   360
         TabIndex        =   10
         Top             =   1560
         Width           =   1935
      End
      Begin VB.OptionButton optVenusQ2A3 
         BackColor       =   &H80000012&
         Caption         =   "Core"
         ForeColor       =   &H00FFFFFF&
         Height          =   495
         Left            =   360
         TabIndex        =   9
         Top             =   2280
         Width           =   1935
      End
      Begin VB.OptionButton optVenusQ2A4 
         BackColor       =   &H80000012&
         Caption         =   "Magnetic Field"
         ForeColor       =   &H00FFFFFF&
         Height          =   495
         Left            =   360
         TabIndex        =   8
         Top             =   3000
         Width           =   1935
      End
      Begin VB.Label lblVenusQuizQuestion3 
         BackColor       =   &H80000012&
         Caption         =   "Venus is lacking a/n _____________ because of it's slow rotation."
         ForeColor       =   &H00FFFFFF&
         Height          =   495
         Left            =   120
         TabIndex        =   12
         Top             =   240
         Width           =   2895
      End
   End
   Begin VB.Frame fraVenusQ4 
      BackColor       =   &H80000012&
      Caption         =   "Question 4"
      ForeColor       =   &H00FFFFFF&
      Height          =   3615
      Left            =   3480
      TabIndex        =   1
      Top             =   3840
      Width           =   3255
      Begin VB.OptionButton optVenusQ4A1 
         BackColor       =   &H80000012&
         Caption         =   "So close to the Sun."
         ForeColor       =   &H00FFFFFF&
         Height          =   495
         Left            =   240
         TabIndex        =   5
         Top             =   840
         Width           =   1935
      End
      Begin VB.OptionButton optVenusQ4A3 
         BackColor       =   &H80000012&
         Caption         =   "Very cloudy."
         ForeColor       =   &H00FFFFFF&
         Height          =   495
         Left            =   240
         TabIndex        =   4
         Top             =   2280
         Width           =   1935
      End
      Begin VB.OptionButton optVenusQ4A2 
         BackColor       =   &H80000012&
         Caption         =   "May contain life."
         ForeColor       =   &H00FFFFFF&
         Height          =   495
         Left            =   240
         TabIndex        =   3
         Top             =   1560
         Width           =   1935
      End
      Begin VB.OptionButton optVenusQ4A4 
         BackColor       =   &H80000012&
         Caption         =   "Very Bright."
         ForeColor       =   &H00FFFFFF&
         Height          =   495
         Left            =   240
         TabIndex        =   2
         Top             =   3000
         Width           =   1935
      End
      Begin VB.Label lblVenusQuizQuestion4 
         BackColor       =   &H80000012&
         Caption         =   "Why were all astonomers alike interested in Venus?"
         ForeColor       =   &H00FFFFFF&
         Height          =   495
         Left            =   120
         TabIndex        =   6
         Top             =   240
         Width           =   3015
      End
   End
   Begin VB.CommandButton cmdVenusQuizRefuel 
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
      TabIndex        =   0
      Top             =   7560
      Width           =   1815
   End
End
Attribute VB_Name = "frmVenusQuiz"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'=======================================================================================================================================================================
'Matthew Lal
'June 5, 2010
'Venus Quiz
'The user refuels to continue the adventure by answer the questions correctly.
'=============================================================================================================================
Option Explicit

Private Sub cmdVenusQuizRefuel_Click()
    'Determines what answers are correct and what answers are not.
    If optVenusQ1A1.Value = True Then
    frmGalaxy3.intScoreQ1 = frmGalaxy3.intScoreQ1 + 1
        Else
        frmGalaxy3.intScoreQ1 = frmGalaxy3.intScoreQ1 - 1
    End If
    
    If optVenusQ2A4.Value = True Then
    frmGalaxy3.intScoreQ2 = frmGalaxy3.intScoreQ2 + 1
        Else
        frmGalaxy3.intScoreQ2 = frmGalaxy3.intScoreQ2 - 1
    End If
    
    If optVenusQ3A3.Value = True Then
    frmGalaxy3.intScoreQ3 = frmGalaxy3.intScoreQ3 + 1
        Else
        frmGalaxy3.intScoreQ3 = frmGalaxy3.intScoreQ3 - 1
    End If
    
    If optVenusQ4A2.Value = True Then
    frmGalaxy3.intScoreQ4 = frmGalaxy3.intScoreQ4 + 1
        Else
        frmGalaxy3.intScoreQ4 = frmGalaxy3.intScoreQ4 - 1
    End If
    
    If optVenusQ5A1.Value = True Then
    frmGalaxy3.intScoreQ5 = frmGalaxy3.intScoreQ5 + 1
        Else
        frmGalaxy3.intScoreQ5 = frmGalaxy3.intScoreQ5 - 1
    End If
    
    'Determines the users overall score for the current quiz.
    frmGalaxy3.intScore = frmGalaxy3.intScoreQ1 + frmGalaxy3.intScoreQ2 + frmGalaxy3.intScoreQ3 + frmGalaxy3.intScoreQ4 + frmGalaxy3.intScoreQ5
    
    'Adds the users score to the fuel.
    frmGalaxy1.intFuel = frmGalaxy1.intFuel + frmGalaxy3.intScore
    
    'Closes the quiz and returns to the galaxy map.
    Unload frmVenusQuiz
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


