VERSION 5.00
Begin VB.Form frmSaturnQuiz 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Galactic Adventures - Refueling On Saturn"
   ClientHeight    =   8415
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   10230
   Icon            =   "frmSaturnQuiz.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "frmSaturnQuiz.frx":1272
   ScaleHeight     =   8415
   ScaleWidth      =   10230
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame fraSaturnQ1 
      BackColor       =   &H80000012&
      Caption         =   "Question1"
      ForeColor       =   &H00FFFFFF&
      Height          =   3615
      Left            =   240
      TabIndex        =   17
      Top             =   240
      Width           =   3135
      Begin VB.OptionButton optSaturnQ1A4 
         BackColor       =   &H80000012&
         Caption         =   "Stars"
         ForeColor       =   &H00FFFFFF&
         Height          =   495
         Left            =   360
         TabIndex        =   21
         Top             =   3000
         Width           =   1935
      End
      Begin VB.OptionButton optSaturnQ1A3 
         BackColor       =   &H80000012&
         Caption         =   "Jupiter"
         ForeColor       =   &H00FFFFFF&
         Height          =   495
         Left            =   360
         TabIndex        =   20
         Top             =   2280
         Width           =   1935
      End
      Begin VB.OptionButton optSaturnQ1A2 
         BackColor       =   &H80000012&
         Caption         =   "Saturn"
         ForeColor       =   &H00FFFFFF&
         Height          =   495
         Left            =   360
         TabIndex        =   19
         Top             =   1560
         Width           =   1935
      End
      Begin VB.OptionButton optSaturnQ1A1 
         BackColor       =   &H00000000&
         Caption         =   "Sun"
         ForeColor       =   &H00FFFFFF&
         Height          =   495
         Left            =   360
         TabIndex        =   18
         Top             =   840
         Width           =   1935
      End
      Begin VB.Label lblSaturnQuizQuestion1 
         BackColor       =   &H00000000&
         Caption         =   "From a distance what twinkles?"
         ForeColor       =   &H00FFFFFF&
         Height          =   495
         Index           =   0
         Left            =   120
         TabIndex        =   22
         Top             =   240
         Width           =   2895
      End
   End
   Begin VB.Frame fraSaturnQ5 
      BackColor       =   &H80000012&
      Caption         =   "Question 5"
      ForeColor       =   &H00FFFFFF&
      Height          =   3615
      Left            =   6840
      TabIndex        =   13
      Top             =   240
      Width           =   3255
      Begin VB.OptionButton optSaturnQ5A2 
         BackColor       =   &H80000012&
         Caption         =   "False"
         ForeColor       =   &H00FFFFFF&
         Height          =   495
         Left            =   240
         TabIndex        =   15
         Top             =   1560
         Width           =   1935
      End
      Begin VB.OptionButton optSaturnQ5A1 
         BackColor       =   &H80000012&
         Caption         =   "True"
         ForeColor       =   &H00FFFFFF&
         Height          =   495
         Left            =   240
         TabIndex        =   14
         Top             =   840
         Width           =   1935
      End
      Begin VB.Label lblSaturnQuizQuestion5 
         BackColor       =   &H00000000&
         Caption         =   "Saturn's ring is solid."
         ForeColor       =   &H00FFFFFF&
         Height          =   495
         Left            =   120
         TabIndex        =   16
         Top             =   240
         Width           =   2655
      End
   End
   Begin VB.Frame fraSaturnQ2 
      BackColor       =   &H80000012&
      Caption         =   "Question 2"
      ForeColor       =   &H00FFFFFF&
      Height          =   3615
      Left            =   240
      TabIndex        =   7
      Top             =   3840
      Width           =   3135
      Begin VB.OptionButton optSaturnQ2A4 
         BackColor       =   &H80000012&
         Caption         =   "December 25, 2000"
         ForeColor       =   &H00FFFFFF&
         Height          =   495
         Left            =   360
         TabIndex        =   11
         Top             =   3000
         Width           =   1935
      End
      Begin VB.OptionButton optSaturnQ2A3 
         BackColor       =   &H80000012&
         Caption         =   "July 4, 1979"
         ForeColor       =   &H00FFFFFF&
         Height          =   495
         Left            =   360
         TabIndex        =   10
         Top             =   2280
         Width           =   1935
      End
      Begin VB.OptionButton optSaturnQ2A2 
         BackColor       =   &H80000012&
         Caption         =   "July 4, 2004"
         ForeColor       =   &H00FFFFFF&
         Height          =   495
         Left            =   360
         TabIndex        =   9
         Top             =   1560
         Width           =   1935
      End
      Begin VB.OptionButton optSaturnQ2A1 
         BackColor       =   &H80000012&
         Caption         =   "July 1, 2004"
         ForeColor       =   &H00FFFFFF&
         Height          =   495
         Left            =   360
         TabIndex        =   8
         Top             =   840
         Width           =   1935
      End
      Begin VB.Label lblSaturnQuizQuestion2 
         BackColor       =   &H80000012&
         Caption         =   "When did the Cassini first arrive at Saturn?"
         ForeColor       =   &H00FFFFFF&
         Height          =   495
         Left            =   120
         TabIndex        =   12
         Top             =   240
         Width           =   2895
      End
   End
   Begin VB.Frame fraSaturnQ4 
      BackColor       =   &H80000012&
      Caption         =   "Question 4"
      ForeColor       =   &H00FFFFFF&
      Height          =   3615
      Left            =   3480
      TabIndex        =   1
      Top             =   3840
      Width           =   3255
      Begin VB.OptionButton optSaturnQ4A4 
         BackColor       =   &H80000012&
         Caption         =   "Pioneer 11"
         ForeColor       =   &H00FFFFFF&
         Height          =   495
         Left            =   240
         TabIndex        =   5
         Top             =   3000
         Width           =   1935
      End
      Begin VB.OptionButton opSaturnQ4A2 
         BackColor       =   &H80000012&
         Caption         =   "Venera 7"
         ForeColor       =   &H00FFFFFF&
         Height          =   495
         Left            =   240
         TabIndex        =   4
         Top             =   1560
         Width           =   1935
      End
      Begin VB.OptionButton optSaturnQ4A3 
         BackColor       =   &H80000012&
         Caption         =   "Pioneer 7"
         ForeColor       =   &H00FFFFFF&
         Height          =   495
         Left            =   240
         TabIndex        =   3
         Top             =   2280
         Width           =   1935
      End
      Begin VB.OptionButton optSaturnQ4A1 
         BackColor       =   &H80000012&
         Caption         =   "Apollo 4"
         ForeColor       =   &H00FFFFFF&
         Height          =   495
         Left            =   240
         TabIndex        =   2
         Top             =   840
         Width           =   1935
      End
      Begin VB.Label lblSaturnQuizQuestion4 
         BackColor       =   &H80000012&
         Caption         =   "Saturn was first visited by what spacecraft?"
         ForeColor       =   &H00FFFFFF&
         Height          =   495
         Left            =   120
         TabIndex        =   6
         Top             =   240
         Width           =   3015
      End
   End
   Begin VB.CommandButton cmdSaturnQuizRefuel 
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
   Begin VB.Frame fraSaturnQ3 
      BackColor       =   &H80000012&
      Caption         =   "Question 3"
      ForeColor       =   &H00FFFFFF&
      Height          =   3615
      Left            =   3480
      TabIndex        =   23
      Top             =   240
      Width           =   3255
      Begin VB.OptionButton optSaturnQ3A1 
         BackColor       =   &H80000012&
         Caption         =   "Uranus"
         ForeColor       =   &H00FFFFFF&
         Height          =   495
         Left            =   360
         TabIndex        =   28
         Top             =   840
         Width           =   1935
      End
      Begin VB.OptionButton optSaturnQ3A4 
         BackColor       =   &H80000012&
         Caption         =   "Saturn"
         ForeColor       =   &H00FFFFFF&
         Height          =   495
         Left            =   360
         TabIndex        =   27
         Top             =   3000
         Width           =   1935
      End
      Begin VB.OptionButton optSaturnQ3A2 
         BackColor       =   &H80000012&
         Caption         =   "Venus"
         ForeColor       =   &H00FFFFFF&
         Height          =   495
         Left            =   360
         TabIndex        =   26
         Top             =   1560
         Width           =   1935
      End
      Begin VB.OptionButton optSaturnQ3A3 
         BackColor       =   &H80000012&
         Caption         =   "Neptune"
         ForeColor       =   &H00FFFFFF&
         Height          =   495
         Left            =   360
         TabIndex        =   25
         Top             =   2280
         Width           =   1935
      End
      Begin VB.Label lblSaturnQuizQuestion3 
         BackColor       =   &H80000012&
         Caption         =   "Which planet doesn't belong?"
         ForeColor       =   &H00FFFFFF&
         Height          =   495
         Left            =   120
         TabIndex        =   24
         Top             =   240
         Width           =   3015
      End
   End
End
Attribute VB_Name = "frmSaturnQuiz"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'=======================================================================================================================================================================
'Matthew Lal
'June 5, 2010
'Saturn Quiz
'The user refuels to continue the adventure by answer the questions correctly.
'==========================================================================================================================
Option Explicit

Private Sub cmdSaturnQuizRefuel_Click()
'Determines what answers are correct and what answers are not.
    If optSaturnQ1A4.Value = True Then
    frmGalaxy3.intScoreQ1 = frmGalaxy3.intScoreQ1 + 1
        Else
        frmGalaxy3.intScoreQ1 = frmGalaxy3.intScoreQ1 - 1
    End If
    
    If optSaturnQ2A2.Value = True Then
    frmGalaxy3.intScoreQ2 = frmGalaxy3.intScoreQ2 + 1
        Else
        frmGalaxy3.intScoreQ2 = frmGalaxy3.intScoreQ2 - 1
    End If
    
    If optSaturnQ3A2.Value = True Then
    frmGalaxy3.intScoreQ3 = frmGalaxy3.intScoreQ3 + 1
        Else
        frmGalaxy3.intScoreQ3 = frmGalaxy3.intScoreQ3 - 1
    End If
    
    If optSaturnQ4A4.Value = True Then
    frmGalaxy3.intScoreQ4 = frmGalaxy3.intScoreQ4 + 1
        Else
        frmGalaxy3.intScoreQ4 = frmGalaxy3.intScoreQ4 - 1
    End If
    
    If optSaturnQ5A2.Value = True Then
    frmGalaxy3.intScoreQ5 = frmGalaxy3.intScoreQ5 + 1
        Else
        frmGalaxy3.intScoreQ5 = frmGalaxy3.intScoreQ5 - 1
    End If
    
    'Determines the users overall score for the current quiz.
    frmGalaxy3.intScore = frmGalaxy3.intScoreQ1 + frmGalaxy3.intScoreQ2 + frmGalaxy3.intScoreQ3 + frmGalaxy3.intScoreQ4 + frmGalaxy3.intScoreQ5
    
    'Adds the users score to the fuel.
    frmGalaxy1.intFuel = frmGalaxy1.intFuel + frmGalaxy3.intScore
      
    'Closes the quiz and returns to the galaxy map.
    Unload frmSaturnQuiz
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

Private Sub lblSaturnQuizQuestion32_Click()

End Sub
