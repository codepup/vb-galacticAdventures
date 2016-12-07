VERSION 5.00
Begin VB.Form frmJupiterQuiz 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Galactic Adventures - Refueling On Jupiter"
   ClientHeight    =   8385
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   10200
   Icon            =   "frmJupiterQuiz.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "frmJupiterQuiz.frx":1272
   ScaleHeight     =   8385
   ScaleWidth      =   10200
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame fraJupiterQ5 
      BackColor       =   &H00000000&
      Caption         =   "Question 5"
      ForeColor       =   &H00FFFFFF&
      Height          =   3615
      Left            =   6840
      TabIndex        =   12
      Top             =   240
      Width           =   3255
      Begin VB.OptionButton optJupiterQ5A1 
         BackColor       =   &H00000000&
         Caption         =   "10-15"
         ForeColor       =   &H00FFFFFF&
         Height          =   495
         Left            =   360
         TabIndex        =   21
         Top             =   840
         Width           =   1935
      End
      Begin VB.OptionButton optJupiterQ5A2 
         BackColor       =   &H00000000&
         Caption         =   "1321"
         ForeColor       =   &H00FFFFFF&
         Height          =   495
         Left            =   360
         TabIndex        =   20
         Top             =   1560
         Width           =   1935
      End
      Begin VB.OptionButton optJupiterQ5A3 
         BackColor       =   &H00000000&
         Caption         =   "400"
         ForeColor       =   &H00FFFFFF&
         Height          =   495
         Left            =   360
         TabIndex        =   19
         Top             =   2280
         Width           =   1935
      End
      Begin VB.OptionButton optJupiterQ5A4 
         BackColor       =   &H00000000&
         Caption         =   "30000"
         ForeColor       =   &H00FFFFFF&
         Height          =   495
         Left            =   360
         TabIndex        =   18
         Top             =   3000
         Width           =   1935
      End
      Begin VB.Label lblJupiterQuizQuestion5 
         BackColor       =   &H00000000&
         Caption         =   "How many Earth's can fit in Jupiter"
         ForeColor       =   &H00FFFFFF&
         Height          =   495
         Left            =   120
         TabIndex        =   37
         Top             =   240
         Width           =   3015
      End
   End
   Begin VB.Frame fraJupiterQ2 
      BackColor       =   &H00000000&
      Caption         =   "Question 2"
      ForeColor       =   &H00FFFFFF&
      Height          =   3615
      Left            =   120
      TabIndex        =   11
      Top             =   3840
      Width           =   3255
      Begin VB.OptionButton optJupiterQ2A1 
         BackColor       =   &H00000000&
         Caption         =   "Size"
         ForeColor       =   &H00FFFFFF&
         Height          =   495
         Left            =   240
         TabIndex        =   25
         Top             =   840
         Width           =   1935
      End
      Begin VB.OptionButton optJupiterQ2A2 
         BackColor       =   &H00000000&
         Caption         =   "Position From The Sun"
         ForeColor       =   &H00FFFFFF&
         Height          =   495
         Left            =   240
         TabIndex        =   24
         Top             =   1560
         Width           =   1935
      End
      Begin VB.OptionButton optJupiterQ2A3 
         BackColor       =   &H00000000&
         Caption         =   "Made Up Of Gas"
         ForeColor       =   &H00FFFFFF&
         Height          =   495
         Left            =   240
         TabIndex        =   23
         Top             =   2280
         Width           =   1935
      End
      Begin VB.OptionButton optJupiterQ2A4 
         BackColor       =   &H00000000&
         Caption         =   "Generates More Radiation Than It Receives"
         ForeColor       =   &H00FFFFFF&
         Height          =   615
         Left            =   240
         TabIndex        =   22
         Top             =   2880
         Width           =   1935
      End
      Begin VB.Label lblJupiterQuizQuestion3 
         BackColor       =   &H00000000&
         Caption         =   "What makes Jupiter unique in the solar system?"
         ForeColor       =   &H00FFFFFF&
         Height          =   495
         Left            =   120
         TabIndex        =   26
         Top             =   240
         Width           =   3015
      End
   End
   Begin VB.Frame fraJupiterQ4 
      BackColor       =   &H00000000&
      Caption         =   "Question 4"
      ForeColor       =   &H00FFFFFF&
      Height          =   3615
      Left            =   3480
      TabIndex        =   10
      Top             =   3840
      Width           =   3255
      Begin VB.OptionButton optJupiterQ4A1 
         BackColor       =   &H00000000&
         Caption         =   "Interrupt Other Planetary Orbits"
         ForeColor       =   &H00FFFFFF&
         Height          =   495
         Left            =   240
         TabIndex        =   30
         Top             =   840
         Width           =   1935
      End
      Begin VB.OptionButton optJupiterQ4A3 
         BackColor       =   &H00000000&
         Caption         =   "Become A Star"
         ForeColor       =   &H00FFFFFF&
         Height          =   495
         Left            =   240
         TabIndex        =   29
         Top             =   2280
         Width           =   1935
      End
      Begin VB.OptionButton optJupiterQ4A2 
         BackColor       =   &H00000000&
         Caption         =   "Shrink"
         ForeColor       =   &H00FFFFFF&
         Height          =   495
         Left            =   240
         TabIndex        =   28
         Top             =   1560
         Width           =   1935
      End
      Begin VB.OptionButton optJupiterQ4A4 
         BackColor       =   &H00000000&
         Caption         =   "It Surface Would Solidate"
         ForeColor       =   &H00FFFFFF&
         Height          =   495
         Left            =   240
         TabIndex        =   27
         Top             =   3000
         Width           =   1935
      End
      Begin VB.Label lblJupiterQuizQuestion4 
         BackColor       =   &H00000000&
         Caption         =   "Jupiter is as big as can be. What would happen if Jupiter got bigger?"
         ForeColor       =   &H00FFFFFF&
         Height          =   495
         Left            =   120
         TabIndex        =   31
         Top             =   240
         Width           =   3015
      End
   End
   Begin VB.Frame fraJupiterQ6 
      BackColor       =   &H00000000&
      Caption         =   "Question 6"
      ForeColor       =   &H00FFFFFF&
      Height          =   3615
      Left            =   6840
      TabIndex        =   9
      Top             =   3840
      Width           =   3255
      Begin VB.OptionButton optJupiterQ6A1 
         BackColor       =   &H00000000&
         Caption         =   "Zones"
         ForeColor       =   &H00FFFFFF&
         Height          =   495
         Left            =   240
         TabIndex        =   35
         Top             =   840
         Width           =   1935
      End
      Begin VB.OptionButton optJupiterQ6A2 
         BackColor       =   &H00000000&
         Caption         =   "Great Red Spot"
         ForeColor       =   &H00FFFFFF&
         Height          =   495
         Left            =   240
         TabIndex        =   34
         Top             =   1560
         Width           =   1935
      End
      Begin VB.OptionButton optJupiterQ6A3 
         BackColor       =   &H00000000&
         Caption         =   "Belts"
         ForeColor       =   &H00FFFFFF&
         Height          =   495
         Left            =   240
         TabIndex        =   33
         Top             =   2280
         Width           =   1935
      End
      Begin VB.OptionButton optJupiterQ6A4 
         BackColor       =   &H00000000&
         Caption         =   "Kelvin"
         ForeColor       =   &H00FFFFFF&
         Height          =   495
         Left            =   240
         TabIndex        =   32
         Top             =   3000
         Width           =   1935
      End
      Begin VB.Label lblJupiterQuizQuestion6 
         BackColor       =   &H00000000&
         Caption         =   "What are the dark bands that swirl around Jupiter called?"
         ForeColor       =   &H00FFFFFF&
         Height          =   495
         Left            =   120
         TabIndex        =   36
         Top             =   240
         Width           =   3015
      End
   End
   Begin VB.Frame fraJupiterQ3 
      BackColor       =   &H00000000&
      Caption         =   "Question 3"
      ForeColor       =   &H00FFFFFF&
      Height          =   3615
      Left            =   3480
      TabIndex        =   8
      Top             =   240
      Width           =   3255
      Begin VB.OptionButton optJupiterQ3A1 
         BackColor       =   &H00000000&
         Caption         =   "Sun"
         ForeColor       =   &H00FFFFFF&
         Height          =   495
         Left            =   240
         TabIndex        =   16
         Top             =   840
         Width           =   1935
      End
      Begin VB.OptionButton optJupiterQ3A4 
         BackColor       =   &H00000000&
         Caption         =   "Jupiter's Core"
         ForeColor       =   &H00FFFFFF&
         Height          =   495
         Left            =   240
         TabIndex        =   15
         Top             =   3000
         Width           =   1935
      End
      Begin VB.OptionButton optJupiterQ3A2 
         BackColor       =   &H00000000&
         Caption         =   "Cooler Temperatures"
         ForeColor       =   &H00FFFFFF&
         Height          =   495
         Left            =   240
         TabIndex        =   14
         Top             =   1560
         Width           =   1935
      End
      Begin VB.OptionButton optJupiterQ3A3 
         BackColor       =   &H00000000&
         Caption         =   "Storm"
         ForeColor       =   &H00FFFFFF&
         Height          =   495
         Left            =   240
         TabIndex        =   13
         Top             =   2280
         Width           =   1935
      End
      Begin VB.Label lblJupiterQuizQuestion2 
         BackColor       =   &H00000000&
         Caption         =   "What is the Great Red Spot?"
         ForeColor       =   &H00FFFFFF&
         Height          =   495
         Left            =   120
         TabIndex        =   17
         Top             =   240
         Width           =   3015
      End
   End
   Begin VB.Frame fraJupiterQ1 
      BackColor       =   &H00000000&
      Caption         =   "Question 1"
      ForeColor       =   &H00FFFFFF&
      Height          =   3615
      Left            =   120
      TabIndex        =   2
      Top             =   240
      Width           =   3255
      Begin VB.OptionButton optJupiterQ1A1 
         BackColor       =   &H00000000&
         Caption         =   "A Mythological Greek God"
         ForeColor       =   &H00FFFFFF&
         Height          =   495
         Left            =   240
         TabIndex        =   6
         Top             =   840
         Width           =   1935
      End
      Begin VB.OptionButton optJupiterQ1A2 
         BackColor       =   &H00000000&
         Caption         =   "Planet"
         ForeColor       =   &H00FFFFFF&
         Height          =   495
         Left            =   240
         TabIndex        =   5
         Top             =   1560
         Width           =   1935
      End
      Begin VB.OptionButton optJupiterQ1A3 
         BackColor       =   &H00000000&
         Caption         =   "Gas Giant"
         ForeColor       =   &H00FFFFFF&
         Height          =   495
         Left            =   240
         TabIndex        =   4
         Top             =   2280
         Width           =   1935
      End
      Begin VB.OptionButton optJupiterQ1A4 
         BackColor       =   &H00000000&
         Caption         =   "A Star"
         ForeColor       =   &H00FFFFFF&
         Height          =   495
         Left            =   240
         TabIndex        =   3
         Top             =   3000
         Width           =   1935
      End
      Begin VB.Label lblJupiterQuizQuestion1 
         BackColor       =   &H00000000&
         Caption         =   "What is Jupiter?"
         ForeColor       =   &H00FFFFFF&
         Height          =   495
         Index           =   0
         Left            =   120
         TabIndex        =   7
         Top             =   240
         Width           =   2895
      End
   End
   Begin VB.CommandButton cmdJupiterQuizRefuel 
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
      TabIndex        =   1
      Top             =   7560
      Width           =   1815
   End
   Begin VB.OptionButton Option1 
      Caption         =   "Option1"
      Height          =   195
      Left            =   10320
      TabIndex        =   0
      Top             =   4560
      Width           =   75
   End
End
Attribute VB_Name = "frmJupiterQuiz"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'=======================================================================================================================================================================
'Matthew Lal
'June 5, 2010
'Jupiter Quiz
'The user refuels to continue the adventure by answer the questions correctly.
'=============================================================================================================================
Option Explicit

Private Sub cmdJupiterQuizRefuel_Click()
'Determines what answers are correct and what answers are not.
 If optJupiterQ1A3.Value = True Then
    frmGalaxy3.intScoreQ1 = frmGalaxy3.intScoreQ1 + 1
        Else
        frmGalaxy3.intScoreQ1 = frmGalaxy3.intScoreQ1 - 1
    End If
    
    If optJupiterQ2A1.Value = True Then
   frmGalaxy3.intScoreQ2 = frmGalaxy3.intScoreQ2 + 1
        Else
        frmGalaxy3.intScoreQ2 = frmGalaxy3.intScoreQ2 - 1
    End If
    
    If optJupiterQ3A3.Value = True Then
    frmGalaxy3.intScoreQ3 = frmGalaxy3.intScoreQ3 + 1
        Else
       frmGalaxy3.intScoreQ3 = frmGalaxy3.intScoreQ3 - 1
    End If
    
    If optJupiterQ4A2.Value = True Then
    frmGalaxy3.intScoreQ4 = frmGalaxy3.intScoreQ4 + 1
        Else
        frmGalaxy3.intScoreQ4 = frmGalaxy3.intScoreQ4 - 1
    End If
    
    If optJupiterQ5A2.Value = True Then
    frmGalaxy3.intScoreQ5 = frmGalaxy3.intScoreQ5 + 1
        Else
        frmGalaxy3.intScoreQ5 = frmGalaxy3.intScoreQ5 - 1
    End If
    If optJupiterQ6A3.Value = True Then
    frmGalaxy3.intScoreQ6 = frmGalaxy3.intScoreQ6 + 1
        Else
        frmGalaxy3.intScoreQ6 = frmGalaxy3.intScoreQ6 - 1
    End If
    
    'Determines the users overall score for the current quiz.
    frmGalaxy3.intScore = frmGalaxy3.intScoreQ1 + frmGalaxy3.intScoreQ2 + frmGalaxy3.intScoreQ3 + frmGalaxy3.intScoreQ4 + frmGalaxy3.intScoreQ5 + frmGalaxy3.intScoreQ6
    
    'Adds the users score to the fuel.
    frmGalaxy1.intFuel = frmGalaxy1.intFuel + frmGalaxy3.intScore
    
    'Closes the quiz and returns to the galaxy map.
    Unload frmJupiterQuiz
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

