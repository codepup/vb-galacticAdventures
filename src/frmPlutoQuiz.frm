VERSION 5.00
Begin VB.Form frmPlutoQuiz 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Galactic Adventures - Refueling On Pluto"
   ClientHeight    =   8385
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   10200
   Icon            =   "frmPlutoQuiz.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "frmPlutoQuiz.frx":1272
   ScaleHeight     =   8385
   ScaleWidth      =   10200
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame fraPlutoQ3 
      BackColor       =   &H00000000&
      Caption         =   "Question 3"
      ForeColor       =   &H00FFFFFF&
      Height          =   3615
      Left            =   3480
      TabIndex        =   25
      Top             =   240
      Width           =   3255
      Begin VB.OptionButton optPlutoQ3A3 
         BackColor       =   &H00000000&
         Caption         =   "Kuiper"
         ForeColor       =   &H00FFFFFF&
         Height          =   495
         Left            =   360
         TabIndex        =   29
         Top             =   2160
         Width           =   1935
      End
      Begin VB.OptionButton optPlutoQ3A2 
         BackColor       =   &H00000000&
         Caption         =   "Kaiser"
         ForeColor       =   &H00FFFFFF&
         Height          =   495
         Left            =   360
         TabIndex        =   28
         Top             =   1440
         Width           =   1935
      End
      Begin VB.OptionButton optPlutoQ3A4 
         BackColor       =   &H00000000&
         Caption         =   "Asteriod"
         ForeColor       =   &H00FFFFFF&
         Height          =   495
         Left            =   360
         TabIndex        =   27
         Top             =   2880
         Width           =   1935
      End
      Begin VB.OptionButton optPlutoQ3A1 
         BackColor       =   &H00000000&
         Caption         =   "Leather"
         ForeColor       =   &H00FFFFFF&
         Height          =   495
         Left            =   360
         TabIndex        =   26
         Top             =   720
         Width           =   1935
      End
      Begin VB.Label lbPlutoQuizQuestion2 
         BackColor       =   &H00000000&
         Caption         =   "Dwarf planets and other objects are found at the __________ belt."
         ForeColor       =   &H00FFFFFF&
         Height          =   495
         Left            =   120
         TabIndex        =   30
         Top             =   240
         Width           =   3015
      End
   End
   Begin VB.Frame fraPlutoQ4 
      BackColor       =   &H00000000&
      Caption         =   "Question 4"
      ForeColor       =   &H00FFFFFF&
      Height          =   3615
      Left            =   3480
      TabIndex        =   4
      Top             =   3960
      Width           =   3255
      Begin VB.OptionButton optPlutoQ4A1 
         BackColor       =   &H00000000&
         Caption         =   "Always small"
         ForeColor       =   &H00FFFFFF&
         Height          =   495
         Left            =   360
         TabIndex        =   8
         Top             =   720
         Width           =   1935
      End
      Begin VB.OptionButton optPlutoQ4A3 
         BackColor       =   &H00000000&
         Caption         =   "A star"
         ForeColor       =   &H00FFFFFF&
         Height          =   495
         Left            =   360
         TabIndex        =   7
         Top             =   2160
         Width           =   1935
      End
      Begin VB.OptionButton optPlutoQ4A2 
         BackColor       =   &H00000000&
         Caption         =   "Once The Size Of The Earth"
         ForeColor       =   &H00FFFFFF&
         Height          =   495
         Left            =   360
         TabIndex        =   6
         Top             =   1440
         Width           =   1935
      End
      Begin VB.OptionButton optPlutoQ4A4 
         BackColor       =   &H00000000&
         Caption         =   "Discovered in 1880"
         ForeColor       =   &H00FFFFFF&
         Height          =   495
         Left            =   360
         TabIndex        =   5
         Top             =   2880
         Width           =   1935
      End
      Begin VB.Label lblPlutoQuizQuestion4 
         BackColor       =   &H00000000&
         Caption         =   "Pluto was..."
         ForeColor       =   &H00FFFFFF&
         Height          =   495
         Left            =   120
         TabIndex        =   9
         Top             =   240
         Width           =   3015
      End
   End
   Begin VB.Frame fraPlutoQ2 
      BackColor       =   &H00000000&
      Caption         =   "Question 3"
      ForeColor       =   &H00FFFFFF&
      Height          =   3615
      Left            =   120
      TabIndex        =   3
      Top             =   3960
      Width           =   3255
      Begin VB.OptionButton optPlutoQ2A1 
         BackColor       =   &H00000000&
         Caption         =   "NASA"
         ForeColor       =   &H00FFFFFF&
         Height          =   495
         Left            =   360
         TabIndex        =   13
         Top             =   720
         Width           =   1935
      End
      Begin VB.OptionButton optPlutoQ2A2 
         BackColor       =   &H00000000&
         Caption         =   "AUS"
         ForeColor       =   &H00FFFFFF&
         Height          =   495
         Left            =   360
         TabIndex        =   12
         Top             =   1440
         Width           =   1935
      End
      Begin VB.OptionButton optPlutoQ2A3 
         BackColor       =   &H00000000&
         Caption         =   "IAU"
         ForeColor       =   &H00FFFFFF&
         Height          =   495
         Left            =   360
         TabIndex        =   11
         Top             =   2160
         Width           =   1935
      End
      Begin VB.OptionButton optPlutoQ2A4 
         BackColor       =   &H00000000&
         Caption         =   "GAU"
         ForeColor       =   &H00FFFFFF&
         Height          =   495
         Left            =   360
         TabIndex        =   10
         Top             =   2880
         Width           =   1935
      End
      Begin VB.Label lblPlutoQuizQuestion3 
         BackColor       =   &H00000000&
         Caption         =   "What group classified Pluto as a dwarf planet?"
         ForeColor       =   &H00FFFFFF&
         Height          =   495
         Left            =   120
         TabIndex        =   14
         Top             =   240
         Width           =   3015
      End
   End
   Begin VB.Frame fraPlutoQ5 
      BackColor       =   &H00000000&
      Caption         =   "Question 5"
      ForeColor       =   &H00FFFFFF&
      Height          =   3615
      Left            =   6840
      TabIndex        =   2
      Top             =   240
      Width           =   3255
      Begin VB.OptionButton optPlutoQ5A1 
         BackColor       =   &H00000000&
         Caption         =   "2012"
         ForeColor       =   &H00FFFFFF&
         Height          =   495
         Left            =   360
         TabIndex        =   18
         Top             =   840
         Width           =   1935
      End
      Begin VB.OptionButton optPlutoQ5A2 
         BackColor       =   &H00000000&
         Caption         =   "2009"
         ForeColor       =   &H00FFFFFF&
         Height          =   495
         Left            =   360
         TabIndex        =   17
         Top             =   1560
         Width           =   1935
      End
      Begin VB.OptionButton optPlutoQ5A3 
         BackColor       =   &H00000000&
         Caption         =   "2015"
         ForeColor       =   &H00FFFFFF&
         Height          =   495
         Left            =   360
         TabIndex        =   16
         Top             =   2280
         Width           =   1935
      End
      Begin VB.OptionButton optPlutoQ5A4 
         BackColor       =   &H00000000&
         Caption         =   "2020"
         ForeColor       =   &H00FFFFFF&
         Height          =   495
         Left            =   360
         TabIndex        =   15
         Top             =   3000
         Width           =   1935
      End
      Begin VB.Label lblPlutoQuizQuestion5 
         BackColor       =   &H00000000&
         Caption         =   "Like all distant planets, Pluto is unknown. More can be learned about it. New Horizons will arrive at Pluto on..."
         ForeColor       =   &H00FFFFFF&
         Height          =   615
         Left            =   120
         TabIndex        =   19
         Top             =   240
         Width           =   3015
      End
   End
   Begin VB.Frame fraPlutoQ1 
      BackColor       =   &H00000000&
      Caption         =   "Question 1"
      ForeColor       =   &H00FFFFFF&
      Height          =   3615
      Left            =   120
      TabIndex        =   1
      Top             =   240
      Width           =   3255
      Begin VB.OptionButton optPlutoQ1A1 
         BackColor       =   &H00000000&
         Caption         =   "1994"
         ForeColor       =   &H00FFFFFF&
         Height          =   495
         Left            =   360
         TabIndex        =   23
         Top             =   840
         Width           =   1935
      End
      Begin VB.OptionButton optPlutoQ1A2 
         BackColor       =   &H00000000&
         Caption         =   "2005"
         ForeColor       =   &H00FFFFFF&
         Height          =   495
         Left            =   360
         TabIndex        =   22
         Top             =   1560
         Width           =   1935
      End
      Begin VB.OptionButton optPlutoQ1A3 
         BackColor       =   &H00000000&
         Caption         =   "2001"
         ForeColor       =   &H00FFFFFF&
         Height          =   495
         Left            =   360
         TabIndex        =   21
         Top             =   2280
         Width           =   1935
      End
      Begin VB.OptionButton optPlutoQ1A4 
         BackColor       =   &H00000000&
         Caption         =   "2006"
         ForeColor       =   &H00FFFFFF&
         Height          =   495
         Left            =   360
         TabIndex        =   20
         Top             =   3000
         Width           =   1935
      End
      Begin VB.Label lblPlutoQuizQuestion1 
         BackColor       =   &H00000000&
         Caption         =   "At what year was Pluto considered a dwarf planet?"
         ForeColor       =   &H00FFFFFF&
         Height          =   495
         Index           =   0
         Left            =   120
         TabIndex        =   24
         Top             =   240
         Width           =   2895
      End
   End
   Begin VB.CommandButton cmdPlutoQuizRefuel 
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
Attribute VB_Name = "frmPlutoQuiz"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'=======================================================================================================================================================================
'Matthew Lal
'June 5, 2010
'Pluto Quiz
'The user refuels to continue the adventure by answer the questions correctly.
'=============================================================================================================================
Option Explicit

Private Sub cmdPlutoQuizRefuel_Click()
    'Determines what answers are correct and what answers are not.
    If optPlutoQ1A4.Value = True Then
    frmGalaxy3.intScoreQ1 = frmGalaxy3.intScoreQ1 + 1
        Else
        frmGalaxy3.intScoreQ1 = frmGalaxy3.intScoreQ1 - 1
    End If
    
    If optPlutoQ2A3.Value = True Then
    frmGalaxy3.intScoreQ2 = frmGalaxy3.intScoreQ2 + 1
        Else
        frmGalaxy3.intScoreQ2 = frmGalaxy3.intScoreQ2 - 1
    End If
    
    If optPlutoQ3A3.Value = True Then
    frmGalaxy3.intScoreQ3 = frmGalaxy3.intScoreQ3 + 1
        Else
        frmGalaxy3.intScoreQ3 = frmGalaxy3.intScoreQ3 - 1
    End If
    
    If optPlutoQ4A2.Value = True Then
    frmGalaxy3.intScoreQ4 = frmGalaxy3.intScoreQ4 + 1
        Else
        frmGalaxy3.intScoreQ4 = frmGalaxy3.intScoreQ4 - 1
    End If
    
    If optPlutoQ5A3.Value = True Then
    frmGalaxy3.intScoreQ5 = frmGalaxy3.intScoreQ5 + 1
        Else
        frmGalaxy3.intScoreQ5 = frmGalaxy3.intScoreQ5 - 1
    End If
    
    'Determines the users overall score for the current quiz.
    frmGalaxy3.intScore = frmGalaxy3.intScoreQ1 + frmGalaxy3.intScoreQ2 + frmGalaxy3.intScoreQ3 + frmGalaxy3.intScoreQ4 + frmGalaxy3.intScoreQ5
    
    'Adds the users score to the fuel.
    frmGalaxy1.intFuel = frmGalaxy1.intFuel + frmGalaxy3.intScore
    
    'Closes the quiz and returns to the galaxy map.
    Unload frmPlutoQuiz
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
