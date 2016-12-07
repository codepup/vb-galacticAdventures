VERSION 5.00
Begin VB.Form frmMercuryQuiz 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Galactic Adventures - Refueling On Mercury"
   ClientHeight    =   8415
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   10230
   Icon            =   "frmMercuryQuiz.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "frmMercuryQuiz.frx":1272
   ScaleHeight     =   8415
   ScaleWidth      =   10230
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame fraMercuryQ1 
      BackColor       =   &H00000000&
      Caption         =   "Question 1"
      ForeColor       =   &H00FFFFFF&
      Height          =   3615
      Left            =   120
      TabIndex        =   23
      Top             =   240
      Width           =   3255
      Begin VB.OptionButton optMercuryQ1A1 
         BackColor       =   &H00000000&
         Caption         =   "True"
         ForeColor       =   &H00FFFFFF&
         Height          =   495
         Left            =   360
         TabIndex        =   25
         Top             =   840
         Width           =   1935
      End
      Begin VB.OptionButton optMercuryQ1A2 
         BackColor       =   &H00000000&
         Caption         =   "False"
         ForeColor       =   &H00FFFFFF&
         Height          =   495
         Left            =   360
         TabIndex        =   24
         Top             =   1560
         Width           =   1935
      End
      Begin VB.Label lblMercuryQuizQuestion1 
         BackColor       =   &H00000000&
         Caption         =   "Mercury has the most extreme tempereature ranges in all the solar system."
         ForeColor       =   &H00FFFFFF&
         Height          =   615
         Index           =   0
         Left            =   120
         TabIndex        =   26
         Top             =   240
         Width           =   2895
      End
   End
   Begin VB.Frame fraMercuryQ4 
      BackColor       =   &H00000000&
      Caption         =   "Question 4"
      ForeColor       =   &H00FFFFFF&
      Height          =   3615
      Left            =   3480
      TabIndex        =   4
      Top             =   3840
      Width           =   3255
      Begin VB.OptionButton optMercuryQ4A1 
         BackColor       =   &H00000000&
         Caption         =   "True"
         ForeColor       =   &H00FFFFFF&
         Height          =   495
         Left            =   360
         TabIndex        =   6
         Top             =   840
         Width           =   1935
      End
      Begin VB.OptionButton optMercuryQ4A2 
         BackColor       =   &H00000000&
         Caption         =   "False"
         ForeColor       =   &H00FFFFFF&
         Height          =   495
         Left            =   360
         TabIndex        =   5
         Top             =   1560
         Width           =   1935
      End
      Begin VB.Label lblMercuryQuizQuestion4 
         BackColor       =   &H00000000&
         Caption         =   "Is Mercury the most dense planet in the solar system?"
         ForeColor       =   &H00FFFFFF&
         Height          =   495
         Left            =   120
         TabIndex        =   7
         Top             =   240
         Width           =   3015
      End
   End
   Begin VB.Frame fraMercuryQ2 
      BackColor       =   &H00000000&
      Caption         =   "Question 2"
      ForeColor       =   &H00FFFFFF&
      Height          =   3615
      Left            =   120
      TabIndex        =   3
      Top             =   3840
      Width           =   3255
      Begin VB.OptionButton optMercuryQ2A1 
         BackColor       =   &H00000000&
         Caption         =   "Parents"
         ForeColor       =   &H00FFFFFF&
         Height          =   495
         Left            =   360
         TabIndex        =   11
         Top             =   840
         Width           =   1935
      End
      Begin VB.OptionButton optMercuryQ2A2 
         BackColor       =   &H00000000&
         Caption         =   "Heat"
         ForeColor       =   &H00FFFFFF&
         Height          =   495
         Left            =   360
         TabIndex        =   10
         Top             =   1560
         Width           =   1935
      End
      Begin VB.OptionButton optMercuryQ2A3 
         BackColor       =   &H00000000&
         Caption         =   "Solar Flare"
         ForeColor       =   &H00FFFFFF&
         Height          =   495
         Left            =   360
         TabIndex        =   9
         Top             =   2280
         Width           =   1935
      End
      Begin VB.OptionButton optMercuryQ2A4 
         BackColor       =   &H00000000&
         Caption         =   "Venus"
         ForeColor       =   &H00FFFFFF&
         Height          =   495
         Left            =   360
         TabIndex        =   8
         Top             =   3000
         Width           =   1935
      End
      Begin VB.Label lblMercuryQuizQuestion3 
         BackColor       =   &H00000000&
         Caption         =   "What is  to blame of Mercury's thin atmospshere?"
         ForeColor       =   &H00FFFFFF&
         Height          =   495
         Left            =   120
         TabIndex        =   12
         Top             =   240
         Width           =   3015
      End
   End
   Begin VB.Frame fraMercuryQ5 
      BackColor       =   &H00000000&
      Caption         =   "Question 5"
      ForeColor       =   &H00FFFFFF&
      Height          =   3615
      Left            =   6840
      TabIndex        =   2
      Top             =   240
      Width           =   3255
      Begin VB.OptionButton optMercuryQ5A1 
         BackColor       =   &H00000000&
         Caption         =   "Zeus"
         ForeColor       =   &H00FFFFFF&
         Height          =   495
         Left            =   360
         TabIndex        =   16
         Top             =   840
         Width           =   1935
      End
      Begin VB.OptionButton optMercuryQ5A2 
         BackColor       =   &H00000000&
         Caption         =   "Hermes"
         ForeColor       =   &H00FFFFFF&
         Height          =   495
         Left            =   360
         TabIndex        =   15
         Top             =   1560
         Width           =   1935
      End
      Begin VB.OptionButton optMercuryQ5A3 
         BackColor       =   &H00000000&
         Caption         =   "Mercury"
         ForeColor       =   &H00FFFFFF&
         Height          =   495
         Left            =   360
         TabIndex        =   14
         Top             =   2280
         Width           =   1935
      End
      Begin VB.OptionButton optMercuryQ5A4 
         BackColor       =   &H00000000&
         Caption         =   "Atlas"
         ForeColor       =   &H00FFFFFF&
         Height          =   495
         Left            =   360
         TabIndex        =   13
         Top             =   3000
         Width           =   1935
      End
      Begin VB.Label lblMercuryQuizQuestion5 
         BackColor       =   &H00000000&
         Caption         =   "What Greek God is Mercury named after?"
         ForeColor       =   &H00FFFFFF&
         Height          =   495
         Left            =   120
         TabIndex        =   17
         Top             =   240
         Width           =   3015
      End
   End
   Begin VB.Frame fraMercuryQ3 
      BackColor       =   &H00000000&
      Caption         =   "Question 3"
      ForeColor       =   &H00FFFFFF&
      Height          =   3615
      Left            =   3480
      TabIndex        =   1
      Top             =   240
      Width           =   3255
      Begin VB.OptionButton optMercuryQ3A1 
         BackColor       =   &H00000000&
         Caption         =   "Venus"
         ForeColor       =   &H00FFFFFF&
         Height          =   495
         Left            =   360
         TabIndex        =   21
         Top             =   840
         Width           =   1935
      End
      Begin VB.OptionButton optMercuryQ3A4 
         BackColor       =   &H00000000&
         Caption         =   "Mercury"
         ForeColor       =   &H00FFFFFF&
         Height          =   495
         Left            =   360
         TabIndex        =   20
         Top             =   3000
         Width           =   1935
      End
      Begin VB.OptionButton optMercuryQ3A2 
         BackColor       =   &H00000000&
         Caption         =   "Sun"
         ForeColor       =   &H00FFFFFF&
         Height          =   495
         Left            =   360
         TabIndex        =   19
         Top             =   1560
         Width           =   1935
      End
      Begin VB.OptionButton optMercuryQ3A3 
         BackColor       =   &H00000000&
         Caption         =   "Pluto"
         ForeColor       =   &H00FFFFFF&
         Height          =   495
         Left            =   360
         TabIndex        =   18
         Top             =   2280
         Width           =   1935
      End
      Begin VB.Label lblMercuryQuizQuestion2 
         BackColor       =   &H00000000&
         Caption         =   "Mercury's position from ______ cause it to be the hottest planet in the solar system."
         ForeColor       =   &H00FFFFFF&
         Height          =   495
         Left            =   120
         TabIndex        =   22
         Top             =   240
         Width           =   3015
      End
   End
   Begin VB.CommandButton cmdMercuryQuizRefuel 
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
Attribute VB_Name = "frmMercuryQuiz"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'=======================================================================================================================================================================
'Matthew Lal
'June 5, 2010
'Mercury Quiz
'The user refuels to continue the adventure by answer the questions correctly.
'=============================================================================================================================
Option Explicit

Private Sub cmdMercuryQuizRefuel_Click()
    'Determines what answers are correct and what answers are not.
    If optMercuryQ1A1.Value = True Then
    frmGalaxy3.intScoreQ1 = frmGalaxy3.intScoreQ1 + 1
        Else
        frmGalaxy3.intScoreQ1 = frmGalaxy3.intScoreQ1 - 1
    End If
    
    If optMercuryQ2A2.Value = True Then
    frmGalaxy3.intScoreQ2 = frmGalaxy3.intScoreQ2 + 1
        Else
        frmGalaxy3.intScoreQ2 = frmGalaxy3.intScoreQ2 - 1
    End If
    
    If optMercuryQ3A2.Value = True Then
    frmGalaxy3.intScoreQ3 = frmGalaxy3.intScoreQ3 + 1
        Else
        frmGalaxy3.intScoreQ3 = frmGalaxy3.intScoreQ3 - 1
    End If
    
    If optMercuryQ4A2.Value = True Then
    frmGalaxy3.intScoreQ4 = frmGalaxy3.intScoreQ4 + 1
        Else
        frmGalaxy3.intScoreQ4 = frmGalaxy3.intScoreQ4 - 1
    End If
    
    If optMercuryQ5A2.Value = True Then
    frmGalaxy3.intScoreQ5 = frmGalaxy3.intScoreQ5 + 1
        Else
        frmGalaxy3.intScoreQ5 = frmGalaxy3.intScoreQ5 - 1
    End If
    
    'Determines the users overall score for the current quiz.
    frmGalaxy3.intScore = frmGalaxy3.intScoreQ1 + frmGalaxy3.intScoreQ2 + frmGalaxy3.intScoreQ3 + frmGalaxy3.intScoreQ4 + frmGalaxy3.intScoreQ5
    
    'Adds the users score to the fuel.
    frmGalaxy1.intFuel = frmGalaxy1.intFuel + frmGalaxy3.intScore
    
    'Closes the quiz and returns to the galaxy map.
    Unload frmMercuryQuiz
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


