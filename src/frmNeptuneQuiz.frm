VERSION 5.00
Begin VB.Form frmNeptuneQuiz 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Galactic Adventures - Refueling On Neptune"
   ClientHeight    =   8415
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   10230
   Icon            =   "frmNeptuneQuiz.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "frmNeptuneQuiz.frx":1272
   ScaleHeight     =   8415
   ScaleWidth      =   10230
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame fraNeptuneQ4 
      BackColor       =   &H00000000&
      Caption         =   "Question 4"
      ForeColor       =   &H00FFFFFF&
      Height          =   3615
      Left            =   3480
      TabIndex        =   5
      Top             =   3960
      Width           =   3255
      Begin VB.OptionButton optNeptuneQ4A1 
         BackColor       =   &H00000000&
         Caption         =   "Great Red Spot"
         ForeColor       =   &H00FFFFFF&
         Height          =   495
         Left            =   360
         TabIndex        =   9
         Top             =   720
         Width           =   1935
      End
      Begin VB.OptionButton optNeptuneQ4A3 
         BackColor       =   &H00000000&
         Caption         =   "Great Dark Spot"
         ForeColor       =   &H00FFFFFF&
         Height          =   495
         Left            =   360
         TabIndex        =   8
         Top             =   2160
         Width           =   1935
      End
      Begin VB.OptionButton optNeptuneQ4A2 
         BackColor       =   &H00000000&
         Caption         =   "Big Blue Spot"
         ForeColor       =   &H00FFFFFF&
         Height          =   495
         Left            =   360
         TabIndex        =   7
         Top             =   1440
         Width           =   1935
      End
      Begin VB.OptionButton optNeptuneQ4A4 
         BackColor       =   &H00000000&
         Caption         =   "Great Blue Spot"
         ForeColor       =   &H00FFFFFF&
         Height          =   495
         Left            =   360
         TabIndex        =   6
         Top             =   2880
         Width           =   1935
      End
      Begin VB.Label lblNeptuneQuizQuestion4 
         BackColor       =   &H00000000&
         Caption         =   "Like Jupiter, Neptune has a big spot on it, what is it called?"
         ForeColor       =   &H00FFFFFF&
         Height          =   495
         Left            =   120
         TabIndex        =   10
         Top             =   240
         Width           =   3015
      End
   End
   Begin VB.Frame fraNeptuneQ2 
      BackColor       =   &H00000000&
      Caption         =   "Question 2"
      ForeColor       =   &H00FFFFFF&
      Height          =   3615
      Left            =   120
      TabIndex        =   4
      Top             =   3960
      Width           =   3255
      Begin VB.OptionButton optNeptuneQ2A1 
         BackColor       =   &H00000000&
         Caption         =   "Voyager 4"
         ForeColor       =   &H00FFFFFF&
         Height          =   495
         Left            =   360
         TabIndex        =   14
         Top             =   720
         Width           =   1935
      End
      Begin VB.OptionButton optNeptuneQ2A2 
         BackColor       =   &H00000000&
         Caption         =   "Apollo 4"
         ForeColor       =   &H00FFFFFF&
         Height          =   495
         Left            =   360
         TabIndex        =   13
         Top             =   1440
         Width           =   1935
      End
      Begin VB.OptionButton optNeptuneQ2A3 
         BackColor       =   &H00000000&
         Caption         =   "Voyager 2"
         ForeColor       =   &H00FFFFFF&
         Height          =   495
         Left            =   360
         TabIndex        =   12
         Top             =   2160
         Width           =   1935
      End
      Begin VB.OptionButton optNeptuneQ2A4 
         BackColor       =   &H00000000&
         Caption         =   "Venera 7"
         ForeColor       =   &H00FFFFFF&
         Height          =   495
         Left            =   360
         TabIndex        =   11
         Top             =   2880
         Width           =   1935
      End
      Begin VB.Label lblNeptuneQuizQuestion3 
         BackColor       =   &H00000000&
         Caption         =   "Neptune was visited once by what spacecraft?"
         ForeColor       =   &H00FFFFFF&
         Height          =   495
         Left            =   120
         TabIndex        =   15
         Top             =   240
         Width           =   3015
      End
   End
   Begin VB.Frame fraNeptuneQ5 
      BackColor       =   &H00000000&
      Caption         =   "Question 5"
      ForeColor       =   &H00FFFFFF&
      Height          =   3615
      Left            =   6840
      TabIndex        =   3
      Top             =   240
      Width           =   3255
      Begin VB.OptionButton optNeptuneQ5A1 
         BackColor       =   &H00000000&
         Caption         =   "Zipper"
         ForeColor       =   &H00FFFFFF&
         Height          =   495
         Left            =   360
         TabIndex        =   19
         Top             =   840
         Width           =   1935
      End
      Begin VB.OptionButton optNeptuneQ5A2 
         BackColor       =   &H00000000&
         Caption         =   "Thors Rath"
         ForeColor       =   &H00FFFFFF&
         Height          =   495
         Left            =   360
         TabIndex        =   18
         Top             =   1560
         Width           =   1935
      End
      Begin VB.OptionButton optNeptuneQ5A3 
         BackColor       =   &H00000000&
         Caption         =   "Scooter"
         ForeColor       =   &H00FFFFFF&
         Height          =   495
         Left            =   360
         TabIndex        =   17
         Top             =   2280
         Width           =   1935
      End
      Begin VB.OptionButton optNeptuneQ5A4 
         BackColor       =   &H00000000&
         Caption         =   "Lambo"
         ForeColor       =   &H00FFFFFF&
         Height          =   495
         Left            =   360
         TabIndex        =   16
         Top             =   3000
         Width           =   1935
      End
      Begin VB.Label lblNeptuneQuizQuestion5 
         BackColor       =   &H00000000&
         Caption         =   "On Neptune there is a small irregulay white cloud that constantly zips around Neptune, what is it called?"
         ForeColor       =   &H00FFFFFF&
         Height          =   615
         Left            =   120
         TabIndex        =   20
         Top             =   240
         Width           =   3015
      End
   End
   Begin VB.Frame fraNeptuneQ3 
      BackColor       =   &H00000000&
      Caption         =   "Question 3"
      ForeColor       =   &H00FFFFFF&
      Height          =   3615
      Left            =   3480
      TabIndex        =   2
      Top             =   240
      Width           =   3255
      Begin VB.OptionButton optNeptuneQ3A1 
         BackColor       =   &H00000000&
         Caption         =   "August 25, 1989"
         ForeColor       =   &H00FFFFFF&
         Height          =   495
         Left            =   360
         TabIndex        =   24
         Top             =   720
         Width           =   1935
      End
      Begin VB.OptionButton optNeptuneQ3A4 
         BackColor       =   &H00000000&
         Caption         =   "August 25, 1990"
         ForeColor       =   &H00FFFFFF&
         Height          =   495
         Left            =   360
         TabIndex        =   23
         Top             =   2880
         Width           =   1935
      End
      Begin VB.OptionButton optNeptuneQ3A2 
         BackColor       =   &H00000000&
         Caption         =   "June 1, 1990"
         ForeColor       =   &H00FFFFFF&
         Height          =   495
         Left            =   360
         TabIndex        =   22
         Top             =   1440
         Width           =   1935
      End
      Begin VB.OptionButton optNeptuneQ3A3 
         BackColor       =   &H00000000&
         Caption         =   "August 23, 1989"
         ForeColor       =   &H00FFFFFF&
         Height          =   495
         Left            =   360
         TabIndex        =   21
         Top             =   2160
         Width           =   1935
      End
      Begin VB.Label lblNeptuneQuizQuestion2 
         BackColor       =   &H00000000&
         Caption         =   "When was Neptune discovered?"
         ForeColor       =   &H00FFFFFF&
         Height          =   495
         Left            =   120
         TabIndex        =   25
         Top             =   240
         Width           =   3015
      End
   End
   Begin VB.Frame fraNeptuneQ1 
      BackColor       =   &H00000000&
      Caption         =   "Question 1"
      ForeColor       =   &H00FFFFFF&
      Height          =   3615
      Left            =   120
      TabIndex        =   1
      Top             =   240
      Width           =   3255
      Begin VB.OptionButton optNeptuneQ1A1 
         BackColor       =   &H00000000&
         Caption         =   "Pluto"
         ForeColor       =   &H00FFFFFF&
         Height          =   495
         Left            =   360
         TabIndex        =   29
         Top             =   720
         Width           =   1935
      End
      Begin VB.OptionButton optNeptuneQ1A2 
         BackColor       =   &H00000000&
         Caption         =   "Saturn"
         ForeColor       =   &H00FFFFFF&
         Height          =   495
         Left            =   360
         TabIndex        =   28
         Top             =   1440
         Width           =   1935
      End
      Begin VB.OptionButton optNeptuneQ1A3 
         BackColor       =   &H00000000&
         Caption         =   "Uranus"
         ForeColor       =   &H00FFFFFF&
         Height          =   495
         Left            =   360
         TabIndex        =   27
         Top             =   2160
         Width           =   1935
      End
      Begin VB.OptionButton optNeptuneQ1A4 
         BackColor       =   &H00000000&
         Caption         =   "Jupiter"
         ForeColor       =   &H00FFFFFF&
         Height          =   495
         Left            =   360
         TabIndex        =   26
         Top             =   2880
         Width           =   1935
      End
      Begin VB.Label lblNeptuneQuizQuestion1 
         BackColor       =   &H00000000&
         Caption         =   "Neptune was discovered after the discovery of what planet?"
         ForeColor       =   &H00FFFFFF&
         Height          =   495
         Index           =   0
         Left            =   120
         TabIndex        =   30
         Top             =   240
         Width           =   2895
      End
   End
   Begin VB.CommandButton cmdNeptuneQuizRefuel 
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
Attribute VB_Name = "frmNeptuneQuiz"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'=======================================================================================================================================================================
'Matthew Lal
'June 5, 2010
'Neptune Quiz
'The user refuels to continue the adventure by answer the questions correctly.
'=============================================================================================================================
Option Explicit

Private Sub cmdNeptuneQuizRefuel_Click()
    'Determines what answers are correct and what answers are not.
    If optNeptuneQ1A3.Value = True Then
    frmGalaxy3.intScoreQ1 = frmGalaxy3.intScoreQ1 + 1
        Else
        frmGalaxy3.intScoreQ1 = frmGalaxy3.intScoreQ1 - 1
    End If
    
    If optNeptuneQ2A3.Value = True Then
    frmGalaxy3.intScoreQ2 = frmGalaxy3.intScoreQ2 + 1
        Else
        frmGalaxy3.intScoreQ2 = frmGalaxy3.intScoreQ2 - 1
    End If
    
    If optNeptuneQ3A1.Value = True Then
    frmGalaxy3.intScoreQ3 = frmGalaxy3.intScoreQ3 + 1
        Else
        frmGalaxy3.intScoreQ3 = frmGalaxy3.intScoreQ3 - 1
    End If
    
    If optNeptuneQ4A3.Value = True Then
    frmGalaxy3.intScoreQ4 = frmGalaxy3.intScoreQ4 + 1
        Else
        frmGalaxy3.intScoreQ4 = frmGalaxy3.intScoreQ4 - 1
    End If
    
    If optNeptuneQ5A3.Value = True Then
    frmGalaxy3.intScoreQ5 = frmGalaxy3.intScoreQ5 + 1
        Else
        frmGalaxy3.intScoreQ5 = frmGalaxy3.intScoreQ5 - 1
    End If
    
    'Determines the users overall score for the current quiz.
    frmGalaxy3.intScore = frmGalaxy3.intScoreQ1 + frmGalaxy3.intScoreQ2 + frmGalaxy3.intScoreQ3 + frmGalaxy3.intScoreQ4 + frmGalaxy3.intScoreQ5
    
    'Adds the users score to the fuel.
    frmGalaxy1.intFuel = frmGalaxy1.intFuel + frmGalaxy3.intScore
    
    'Closes the quiz and returns to the galaxy map.
    Unload frmNeptuneQuiz
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


