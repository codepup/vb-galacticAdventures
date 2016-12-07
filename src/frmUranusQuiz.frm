VERSION 5.00
Begin VB.Form frmUranusQuiz 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Galactic Adventures - Refueling On Uranus"
   ClientHeight    =   8415
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   10230
   Icon            =   "frmUranusQuiz.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "frmUranusQuiz.frx":1272
   ScaleHeight     =   8415
   ScaleWidth      =   10230
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame fraUranusQ3 
      BackColor       =   &H80000012&
      Caption         =   "Question 3"
      ForeColor       =   &H00FFFFFF&
      Height          =   3615
      Left            =   3480
      TabIndex        =   23
      Top             =   240
      Width           =   3255
      Begin VB.OptionButton optUranusQ3A3 
         BackColor       =   &H80000012&
         Caption         =   "It's made of gas."
         ForeColor       =   &H00FFFFFF&
         Height          =   495
         Left            =   360
         TabIndex        =   27
         Top             =   2280
         Width           =   1935
      End
      Begin VB.OptionButton optUranusQ3A2 
         BackColor       =   &H80000012&
         Caption         =   "It has a ring."
         ForeColor       =   &H00FFFFFF&
         Height          =   495
         Left            =   360
         TabIndex        =   26
         Top             =   1560
         Width           =   1935
      End
      Begin VB.OptionButton optUranusQ3A4 
         BackColor       =   &H80000012&
         Caption         =   "It is blue."
         ForeColor       =   &H00FFFFFF&
         Height          =   495
         Left            =   360
         TabIndex        =   25
         Top             =   3000
         Width           =   1935
      End
      Begin VB.OptionButton optUranusQ3A1 
         BackColor       =   &H80000012&
         Caption         =   "It's side ways."
         ForeColor       =   &H00FFFFFF&
         Height          =   495
         Left            =   360
         TabIndex        =   24
         Top             =   840
         Width           =   1935
      End
      Begin VB.Label lblUranusQuizQuestion2 
         BackColor       =   &H80000012&
         Caption         =   "Uranus is different in what way?"
         ForeColor       =   &H00FFFFFF&
         Height          =   495
         Left            =   120
         TabIndex        =   28
         Top             =   240
         Width           =   3015
      End
   End
   Begin VB.Frame fraUranusQ1 
      BackColor       =   &H80000012&
      Caption         =   "Question1"
      ForeColor       =   &H00FFFFFF&
      Height          =   3615
      Left            =   240
      TabIndex        =   17
      Top             =   240
      Width           =   3135
      Begin VB.OptionButton optUranusQ1A1 
         BackColor       =   &H00000000&
         Caption         =   "Your Butt"
         ForeColor       =   &H00FFFFFF&
         Height          =   495
         Left            =   360
         TabIndex        =   21
         Top             =   840
         Width           =   1935
      End
      Begin VB.OptionButton optUranusQ1A2 
         BackColor       =   &H80000012&
         Caption         =   "Urine Us"
         ForeColor       =   &H00FFFFFF&
         Height          =   495
         Left            =   360
         TabIndex        =   20
         Top             =   1560
         Width           =   1935
      End
      Begin VB.OptionButton optUranusQ1A3 
         BackColor       =   &H80000012&
         Caption         =   "Your Anus"
         ForeColor       =   &H00FFFFFF&
         Height          =   495
         Left            =   360
         TabIndex        =   19
         Top             =   2280
         Width           =   1935
      End
      Begin VB.OptionButton optUranusQ1A4 
         BackColor       =   &H80000012&
         Caption         =   "Yoor A Nus"
         ForeColor       =   &H00FFFFFF&
         Height          =   495
         Left            =   360
         TabIndex        =   18
         Top             =   3000
         Width           =   1935
      End
      Begin VB.Label lblUranusQuizQuestion1 
         BackColor       =   &H00000000&
         Caption         =   "Uranus is pronounced:"
         ForeColor       =   &H00FFFFFF&
         Height          =   495
         Index           =   0
         Left            =   120
         TabIndex        =   22
         Top             =   240
         Width           =   2895
      End
   End
   Begin VB.Frame fraUranusQ5 
      BackColor       =   &H80000012&
      Caption         =   "Question 5"
      ForeColor       =   &H00FFFFFF&
      Height          =   3615
      Left            =   6840
      TabIndex        =   13
      Top             =   240
      Width           =   3255
      Begin VB.OptionButton optUranusQ5A1 
         BackColor       =   &H80000012&
         Caption         =   "True"
         ForeColor       =   &H00FFFFFF&
         Height          =   495
         Left            =   240
         TabIndex        =   15
         Top             =   840
         Width           =   1935
      End
      Begin VB.OptionButton optUranusQ5A2 
         BackColor       =   &H80000012&
         Caption         =   "False"
         ForeColor       =   &H00FFFFFF&
         Height          =   495
         Left            =   240
         TabIndex        =   14
         Top             =   1560
         Width           =   1935
      End
      Begin VB.Label lblUranusQuizQuestion5 
         BackColor       =   &H80000012&
         Caption         =   "All planets, with the exception of Earth, do not have seasonal changes."
         ForeColor       =   &H00FFFFFF&
         Height          =   615
         Left            =   120
         TabIndex        =   16
         Top             =   240
         Width           =   2655
      End
   End
   Begin VB.Frame fraUranusQ2 
      BackColor       =   &H80000012&
      Caption         =   "Question 2"
      ForeColor       =   &H00FFFFFF&
      Height          =   3615
      Left            =   240
      TabIndex        =   7
      Top             =   3840
      Width           =   3135
      Begin VB.OptionButton optUranusQ2A1 
         BackColor       =   &H80000012&
         Caption         =   "William Shakespeare"
         ForeColor       =   &H00FFFFFF&
         Height          =   495
         Left            =   360
         TabIndex        =   11
         Top             =   840
         Width           =   1935
      End
      Begin VB.OptionButton optUranusQ2A2 
         BackColor       =   &H80000012&
         Caption         =   "Will I Am"
         ForeColor       =   &H00FFFFFF&
         Height          =   495
         Left            =   360
         TabIndex        =   10
         Top             =   1560
         Width           =   1935
      End
      Begin VB.OptionButton optUranusQ2A3 
         BackColor       =   &H80000012&
         Caption         =   "William Herschel"
         ForeColor       =   &H00FFFFFF&
         Height          =   495
         Left            =   360
         TabIndex        =   9
         Top             =   2280
         Width           =   1935
      End
      Begin VB.OptionButton optUranusQ2A4 
         BackColor       =   &H80000012&
         Caption         =   "William Kuiper"
         ForeColor       =   &H00FFFFFF&
         Height          =   495
         Left            =   360
         TabIndex        =   8
         Top             =   3000
         Width           =   1935
      End
      Begin VB.Label lblUranusQuizQuestion3 
         BackColor       =   &H80000012&
         Caption         =   "Uranus was discovered by who?"
         ForeColor       =   &H00FFFFFF&
         Height          =   495
         Left            =   120
         TabIndex        =   12
         Top             =   240
         Width           =   2895
      End
   End
   Begin VB.Frame fraUranusQ4 
      BackColor       =   &H80000012&
      Caption         =   "Question 4"
      ForeColor       =   &H00FFFFFF&
      Height          =   3615
      Left            =   3480
      TabIndex        =   1
      Top             =   3840
      Width           =   3255
      Begin VB.OptionButton optUranusQ4A1 
         BackColor       =   &H80000012&
         Caption         =   "An Organ"
         ForeColor       =   &H00FFFFFF&
         Height          =   495
         Left            =   240
         TabIndex        =   5
         Top             =   840
         Width           =   1935
      End
      Begin VB.OptionButton optUranusQ4A3 
         BackColor       =   &H80000012&
         Caption         =   "Dwarf Planet"
         ForeColor       =   &H00FFFFFF&
         Height          =   495
         Left            =   240
         TabIndex        =   4
         Top             =   2280
         Width           =   1935
      End
      Begin VB.OptionButton optUranusQ4A2 
         BackColor       =   &H80000012&
         Caption         =   "Gas Giant"
         ForeColor       =   &H00FFFFFF&
         Height          =   495
         Left            =   240
         TabIndex        =   3
         Top             =   1560
         Width           =   1935
      End
      Begin VB.OptionButton optUranusQ4A4 
         BackColor       =   &H80000012&
         Caption         =   "A Star"
         ForeColor       =   &H00FFFFFF&
         Height          =   495
         Left            =   240
         TabIndex        =   2
         Top             =   3000
         Width           =   1935
      End
      Begin VB.Label lblUranusQuizQuestion4 
         BackColor       =   &H80000012&
         Caption         =   "What kind of planet is Uranus?"
         ForeColor       =   &H00FFFFFF&
         Height          =   495
         Left            =   120
         TabIndex        =   6
         Top             =   240
         Width           =   3015
      End
   End
   Begin VB.CommandButton cmdUranusQuizRefuel 
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
Attribute VB_Name = "frmUranusQuiz"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'=======================================================================================================================================================================
'Matthew Lal
'June 5, 2010
'Uranus Quiz
'The user refuels to continue the adventure by answer the questions correctly.
'=============================================================================================================================
Option Explicit

Private Sub cmdUranusQuizRefuel_Click()
    'Determines what answers are correct and what answers are not.
    If optUranusQ1A4.Value = True Then
    frmGalaxy3.intScoreQ1 = frmGalaxy3.intScoreQ1 + 1
        Else
        frmGalaxy3.intScoreQ1 = frmGalaxy3.intScoreQ1 - 1
    End If
    
    If optUranusQ2A3.Value = True Then
    frmGalaxy3.intScoreQ2 = frmGalaxy3.intScoreQ2 + 1
        Else
        frmGalaxy3.intScoreQ2 = frmGalaxy3.intScoreQ2 - 1
    End If
    
    If optUranusQ3A1.Value = True Then
    frmGalaxy3.intScoreQ3 = frmGalaxy3.intScoreQ3 + 1
        Else
        frmGalaxy3.intScoreQ3 = frmGalaxy3.intScoreQ3 - 1
    End If
    
    If optUranusQ4A2.Value = True Then
    frmGalaxy3.intScoreQ4 = frmGalaxy3.intScoreQ4 + 1
        Else
        frmGalaxy3.intScoreQ4 = frmGalaxy3.intScoreQ4 - 1
    End If
    
    If optUranusQ5A2.Value = True Then
    frmGalaxy3.intScoreQ5 = frmGalaxy3.intScoreQ5 + 1
        Else
        frmGalaxy3.intScoreQ5 = frmGalaxy3.intScoreQ5 - 1
    End If
    
    'Determines the users overall score for the current quiz.
    frmGalaxy3.intScore = frmGalaxy3.intScoreQ1 + frmGalaxy3.intScoreQ2 + frmGalaxy3.intScoreQ3 + frmGalaxy3.intScoreQ4 + frmGalaxy3.intScoreQ5
    
    'Adds the users score to the fuel.
    frmGalaxy1.intFuel = frmGalaxy1.intFuel + frmGalaxy3.intScore
    
    'Closes the quiz and returns to the galaxy map.
    Unload frmUranusQuiz
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
