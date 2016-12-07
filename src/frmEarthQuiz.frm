VERSION 5.00
Begin VB.Form frmEarthQuiz 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Galactic Adventures - Refueling On Earth"
   ClientHeight    =   8385
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   10200
   Icon            =   "frmEarthQuiz.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "frmEarthQuiz.frx":1272
   ScaleHeight     =   8385
   ScaleWidth      =   10200
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame fraEarthQ3 
      BackColor       =   &H00000000&
      Caption         =   "Question 3"
      ForeColor       =   &H00FFFFFF&
      Height          =   3615
      Left            =   3480
      TabIndex        =   11
      Top             =   240
      Width           =   3255
      Begin VB.OptionButton optEarthQ3A1 
         BackColor       =   &H00000000&
         Caption         =   "Oxygen"
         ForeColor       =   &H00FFFFFF&
         Height          =   495
         Left            =   120
         TabIndex        =   15
         Top             =   840
         Width           =   1935
      End
      Begin VB.OptionButton optEarthQ3A4 
         BackColor       =   &H00000000&
         Caption         =   "Carbon Dioxide"
         ForeColor       =   &H00FFFFFF&
         Height          =   495
         Left            =   120
         TabIndex        =   14
         Top             =   3000
         Width           =   1935
      End
      Begin VB.OptionButton optEarthQ3A2 
         BackColor       =   &H00000000&
         Caption         =   "Water Vapour"
         ForeColor       =   &H00FFFFFF&
         Height          =   495
         Left            =   120
         TabIndex        =   13
         Top             =   1560
         Width           =   1935
      End
      Begin VB.OptionButton optEarthQ3A3 
         BackColor       =   &H00000000&
         Caption         =   "Argon"
         ForeColor       =   &H00FFFFFF&
         Height          =   495
         Left            =   120
         TabIndex        =   12
         Top             =   2280
         Width           =   1935
      End
      Begin VB.Label lblEarthQuizQuestion2 
         BackColor       =   &H00000000&
         Caption         =   "Earth's low levels of ________ in the atmosphere keeps the temperature stable."
         ForeColor       =   &H00FFFFFF&
         Height          =   495
         Left            =   120
         TabIndex        =   16
         Top             =   240
         Width           =   3015
      End
   End
   Begin VB.Frame fraEarthQ5 
      BackColor       =   &H00000000&
      Caption         =   "Question 5"
      ForeColor       =   &H00FFFFFF&
      Height          =   3615
      Left            =   6840
      TabIndex        =   10
      Top             =   240
      Width           =   3255
      Begin VB.OptionButton optEarthQ5A1 
         BackColor       =   &H00000000&
         Caption         =   "Oxygen"
         ForeColor       =   &H00FFFFFF&
         Height          =   495
         Left            =   120
         TabIndex        =   20
         Top             =   840
         Width           =   1935
      End
      Begin VB.OptionButton optEarthQ5A2 
         BackColor       =   &H00000000&
         Caption         =   "Water Vapour"
         ForeColor       =   &H00FFFFFF&
         Height          =   495
         Left            =   120
         TabIndex        =   19
         Top             =   1560
         Width           =   1935
      End
      Begin VB.OptionButton optEarthQ5A3 
         BackColor       =   &H00000000&
         Caption         =   "Argon"
         ForeColor       =   &H00FFFFFF&
         Height          =   495
         Left            =   120
         TabIndex        =   18
         Top             =   2280
         Width           =   1935
      End
      Begin VB.OptionButton optEarthQ5A4 
         BackColor       =   &H00000000&
         Caption         =   "Carbon Dioxide"
         ForeColor       =   &H00FFFFFF&
         Height          =   495
         Left            =   120
         TabIndex        =   17
         Top             =   3000
         Width           =   1935
      End
      Begin VB.Label lblEarthQuizQuestion5 
         BackColor       =   &H00000000&
         Caption         =   "What gas is in the atmosphere that is able to support life?"
         ForeColor       =   &H00FFFFFF&
         Height          =   495
         Left            =   120
         TabIndex        =   21
         Top             =   240
         Width           =   3015
      End
   End
   Begin VB.Frame fraEarthQ2 
      BackColor       =   &H00000000&
      Caption         =   "Question 2"
      ForeColor       =   &H00FFFFFF&
      Height          =   3615
      Left            =   120
      TabIndex        =   9
      Top             =   3840
      Width           =   3255
      Begin VB.OptionButton optEarthQ2A1 
         BackColor       =   &H00000000&
         Caption         =   "Atmosphere"
         ForeColor       =   &H00FFFFFF&
         Height          =   495
         Left            =   360
         TabIndex        =   25
         Top             =   840
         Width           =   1935
      End
      Begin VB.OptionButton optEarthQ2A2 
         BackColor       =   &H00000000&
         Caption         =   "Sun"
         ForeColor       =   &H00FFFFFF&
         Height          =   495
         Left            =   360
         TabIndex        =   24
         Top             =   1560
         Width           =   1935
      End
      Begin VB.OptionButton optEarthQ2A3 
         BackColor       =   &H00000000&
         Caption         =   "Pangaea"
         ForeColor       =   &H00FFFFFF&
         Height          =   495
         Left            =   360
         TabIndex        =   23
         Top             =   2280
         Width           =   1935
      End
      Begin VB.OptionButton optEarthQ2A4 
         BackColor       =   &H00000000&
         Caption         =   "Venus"
         ForeColor       =   &H00FFFFFF&
         Height          =   495
         Left            =   360
         TabIndex        =   22
         Top             =   3000
         Width           =   1935
      End
      Begin VB.Label lblEarthQuizQuestion3 
         BackColor       =   &H00000000&
         Caption         =   "Earth's position from _____ plays a vital role in sustaining complex life."
         ForeColor       =   &H00FFFFFF&
         Height          =   495
         Left            =   120
         TabIndex        =   26
         Top             =   240
         Width           =   3015
      End
   End
   Begin VB.Frame fraEarthQ4 
      BackColor       =   &H00000000&
      Caption         =   "Question 4"
      ForeColor       =   &H00FFFFFF&
      Height          =   3615
      Left            =   3480
      TabIndex        =   8
      Top             =   3840
      Width           =   3255
      Begin VB.OptionButton optEarthQ4A1 
         BackColor       =   &H00000000&
         Caption         =   "Magnetic Field"
         ForeColor       =   &H00FFFFFF&
         Height          =   495
         Left            =   120
         TabIndex        =   30
         Top             =   840
         Width           =   1935
      End
      Begin VB.OptionButton optEarthQ4A3 
         BackColor       =   &H00000000&
         Caption         =   "Earth's Core"
         ForeColor       =   &H00FFFFFF&
         Height          =   495
         Left            =   120
         TabIndex        =   29
         Top             =   2280
         Width           =   1935
      End
      Begin VB.OptionButton optEarthQ4A2 
         BackColor       =   &H00000000&
         Caption         =   "Earth's Poles"
         ForeColor       =   &H00FFFFFF&
         Height          =   495
         Left            =   120
         TabIndex        =   28
         Top             =   1560
         Width           =   1935
      End
      Begin VB.OptionButton optEarthQ4A4 
         BackColor       =   &H00000000&
         Caption         =   "Greenhouse Effect"
         ForeColor       =   &H00FFFFFF&
         Height          =   495
         Left            =   120
         TabIndex        =   27
         Top             =   3000
         Width           =   1935
      End
      Begin VB.Label lblEarthQuizQuestion4 
         BackColor       =   &H00000000&
         Caption         =   "How does Earth defend itself from solar flares?"
         ForeColor       =   &H00FFFFFF&
         Height          =   495
         Left            =   120
         TabIndex        =   31
         Top             =   240
         Width           =   3015
      End
   End
   Begin VB.Frame fraEarthQ1 
      BackColor       =   &H00000000&
      Caption         =   "Question 1"
      ForeColor       =   &H00FFFFFF&
      Height          =   3615
      Left            =   120
      TabIndex        =   2
      Top             =   240
      Width           =   3255
      Begin VB.OptionButton optEarthQ1A1 
         BackColor       =   &H00000000&
         Caption         =   "Rocks"
         ForeColor       =   &H00FFFFFF&
         Height          =   495
         Left            =   240
         TabIndex        =   6
         Top             =   840
         Width           =   1935
      End
      Begin VB.OptionButton optEarthQ1A2 
         BackColor       =   &H00000000&
         Caption         =   "Deity"
         ForeColor       =   &H00FFFFFF&
         Height          =   495
         Left            =   240
         TabIndex        =   5
         Top             =   1560
         Width           =   1935
      End
      Begin VB.OptionButton optEarthQ1A3 
         BackColor       =   &H00000000&
         Caption         =   "Water"
         ForeColor       =   &H00FFFFFF&
         Height          =   495
         Left            =   240
         TabIndex        =   4
         Top             =   2280
         Width           =   1935
      End
      Begin VB.OptionButton optEarthQ1A4 
         BackColor       =   &H00000000&
         Caption         =   "Nitrogen"
         ForeColor       =   &H00FFFFFF&
         Height          =   495
         Left            =   240
         TabIndex        =   3
         Top             =   3000
         Width           =   1935
      End
      Begin VB.Label lblEarthQuizQuestion1 
         BackColor       =   &H00000000&
         Caption         =   "What is the is the basic component required to support complex life?"
         ForeColor       =   &H00FFFFFF&
         Height          =   495
         Index           =   0
         Left            =   120
         TabIndex        =   7
         Top             =   240
         Width           =   2895
      End
   End
   Begin VB.CommandButton cmdEarthQuizRefuel 
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
Attribute VB_Name = "frmEarthQuiz"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'=======================================================================================================================================================================
'Matthew Lal
'June 5, 2010
'Earth Quiz
'The user refuels to contunie the adventure by answer the questions correctly.
'=======================================================================================================================================================================
Option Explicit


Private Sub cmdEarthQuizRefuel_Click()

'Determines what answers are correct and what answers are not.
If optEarthQ1A3.Value = True Then
    frmGalaxy3.intScoreQ1 = frmGalaxy3.intScoreQ1 + 1
        Else
        frmGalaxy3.intScoreQ1 = frmGalaxy3.intScoreQ1 - 1
    End If
    
    If optEarthQ2A2.Value = True Then
    frmGalaxy3.intScoreQ2 = frmGalaxy3.intScoreQ2 + 1
        Else
        frmGalaxy3.intScoreQ2 = frmGalaxy3.intScoreQ2 - 1
    End If
    
    If optEarthQ3A4.Value = True Then
    frmGalaxy3.intScoreQ3 = frmGalaxy3.intScoreQ3 + 1
        Else
        frmGalaxy3.intScoreQ3 = frmGalaxy3.intScoreQ3 - 1
    End If
    
    If optEarthQ4A1.Value = True Then
    frmGalaxy3.intScoreQ4 = frmGalaxy3.intScoreQ4 + 1
        Else
        frmGalaxy3.intScoreQ4 = frmGalaxy3.intScoreQ4 - 1
    End If
    
    If optEarthQ5A1.Value = True Then
    frmGalaxy3.intScoreQ5 = frmGalaxy3.intScoreQ5 + 1
        Else
        frmGalaxy3.intScoreQ5 = frmGalaxy3.intScoreQ5 - 1
End If
    
    'Determines the users overall score for the current quiz.
    frmGalaxy3.intScore = frmGalaxy3.intScoreQ1 + frmGalaxy3.intScoreQ2 + frmGalaxy3.intScoreQ3 + frmGalaxy3.intScoreQ4 + frmGalaxy3.intScoreQ5
    
    'Adds the users score to the fuel.
    frmGalaxy1.intFuel = frmGalaxy1.intFuel + frmGalaxy3.intScore
    
    'Closes the quiz and returns to the galaxy map.
    Unload frmEarthQuiz
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


