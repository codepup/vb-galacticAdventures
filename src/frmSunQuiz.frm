VERSION 5.00
Begin VB.Form frmSunQuiz 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Galactic Adventures - Refueling On The Sun"
   ClientHeight    =   8385
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   10200
   Icon            =   "frmSunQuiz.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "frmSunQuiz.frx":1272
   ScaleHeight     =   8385
   ScaleWidth      =   10200
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame fraSunQ3 
      BackColor       =   &H80000012&
      Caption         =   "Question 3"
      ForeColor       =   &H00FFFFFF&
      Height          =   3615
      Left            =   3480
      TabIndex        =   23
      Top             =   240
      Width           =   3255
      Begin VB.OptionButton optSunQ3A3 
         BackColor       =   &H80000012&
         Caption         =   "Light"
         ForeColor       =   &H00FFFFFF&
         Height          =   495
         Left            =   360
         TabIndex        =   27
         Top             =   2280
         Width           =   1935
      End
      Begin VB.OptionButton optSunQ3A2 
         BackColor       =   &H80000012&
         Caption         =   "Arcs"
         ForeColor       =   &H00FFFFFF&
         Height          =   495
         Left            =   360
         TabIndex        =   26
         Top             =   1560
         Width           =   1935
      End
      Begin VB.OptionButton optSunQ3A4 
         BackColor       =   &H80000012&
         Caption         =   "Solar winds"
         ForeColor       =   &H00FFFFFF&
         Height          =   495
         Left            =   360
         TabIndex        =   25
         Top             =   3000
         Width           =   1935
      End
      Begin VB.OptionButton optSunQ3A1 
         BackColor       =   &H80000012&
         Caption         =   "Sun spots"
         ForeColor       =   &H00FFFFFF&
         Height          =   495
         Left            =   360
         TabIndex        =   24
         Top             =   840
         Width           =   1935
      End
      Begin VB.Label lbSunQuizQuestion2 
         BackColor       =   &H80000012&
         Caption         =   "As well as energy the Sun emits________."
         ForeColor       =   &H00FFFFFF&
         Height          =   495
         Left            =   120
         TabIndex        =   28
         Top             =   240
         Width           =   3015
      End
   End
   Begin VB.Frame fraSunQ1 
      BackColor       =   &H80000012&
      Caption         =   "Question1"
      ForeColor       =   &H00FFFFFF&
      Height          =   3615
      Left            =   240
      TabIndex        =   17
      Top             =   240
      Width           =   3135
      Begin VB.OptionButton optSunQ1A1 
         BackColor       =   &H00000000&
         Caption         =   "3"
         ForeColor       =   &H00FFFFFF&
         Height          =   495
         Left            =   360
         TabIndex        =   21
         Top             =   840
         Width           =   1935
      End
      Begin VB.OptionButton optSunQ1A2 
         BackColor       =   &H80000012&
         Caption         =   "5"
         ForeColor       =   &H00FFFFFF&
         Height          =   495
         Left            =   360
         TabIndex        =   20
         Top             =   1560
         Width           =   1935
      End
      Begin VB.OptionButton optSunQ1A3 
         BackColor       =   &H80000012&
         Caption         =   "7"
         ForeColor       =   &H00FFFFFF&
         Height          =   495
         Left            =   360
         TabIndex        =   19
         Top             =   2280
         Width           =   1935
      End
      Begin VB.OptionButton optSunQ1A4 
         BackColor       =   &H80000012&
         Caption         =   "6"
         ForeColor       =   &H00FFFFFF&
         Height          =   495
         Left            =   360
         TabIndex        =   18
         Top             =   3000
         Width           =   1935
      End
      Begin VB.Label lblSunQuizQuestion1 
         BackColor       =   &H00000000&
         Caption         =   "The Sun will eventually run out of fuel and blow up in the next ____ billion years."
         ForeColor       =   &H00FFFFFF&
         Height          =   615
         Index           =   0
         Left            =   120
         TabIndex        =   22
         Top             =   240
         Width           =   2895
      End
   End
   Begin VB.Frame fraSunQ5 
      BackColor       =   &H80000012&
      Caption         =   "Question 5"
      ForeColor       =   &H00FFFFFF&
      Height          =   3615
      Left            =   6840
      TabIndex        =   13
      Top             =   240
      Width           =   3255
      Begin VB.OptionButton optSunQ5A1 
         BackColor       =   &H80000012&
         Caption         =   "True"
         ForeColor       =   &H00FFFFFF&
         Height          =   495
         Left            =   240
         TabIndex        =   15
         Top             =   840
         Width           =   1935
      End
      Begin VB.OptionButton optSunQ5A2 
         BackColor       =   &H80000012&
         Caption         =   "False"
         ForeColor       =   &H00FFFFFF&
         Height          =   495
         Left            =   240
         TabIndex        =   14
         Top             =   1560
         Width           =   1935
      End
      Begin VB.Label lblSunQuizQuestion5 
         BackColor       =   &H80000012&
         Caption         =   "The Sun is a star."
         ForeColor       =   &H00FFFFFF&
         Height          =   495
         Left            =   120
         TabIndex        =   16
         Top             =   240
         Width           =   2655
      End
   End
   Begin VB.Frame fraSunQ2 
      BackColor       =   &H80000012&
      Caption         =   "Question 2"
      ForeColor       =   &H00FFFFFF&
      Height          =   3615
      Left            =   240
      TabIndex        =   7
      Top             =   3840
      Width           =   3135
      Begin VB.OptionButton optSunQ2A1 
         BackColor       =   &H80000012&
         Caption         =   "Nuclear Transfusion"
         ForeColor       =   &H00FFFFFF&
         Height          =   495
         Left            =   360
         TabIndex        =   11
         Top             =   840
         Width           =   1935
      End
      Begin VB.OptionButton optSunQ2A2 
         BackColor       =   &H80000012&
         Caption         =   "Nuclear Fusion"
         ForeColor       =   &H00FFFFFF&
         Height          =   495
         Left            =   360
         TabIndex        =   10
         Top             =   1560
         Width           =   1935
      End
      Begin VB.OptionButton optSunQ2A3 
         BackColor       =   &H80000012&
         Caption         =   "Nuclear Fission"
         ForeColor       =   &H00FFFFFF&
         Height          =   495
         Left            =   360
         TabIndex        =   9
         Top             =   2280
         Width           =   1935
      End
      Begin VB.OptionButton optSunQ2A4 
         BackColor       =   &H80000012&
         Caption         =   "Shake and Bake"
         ForeColor       =   &H00FFFFFF&
         Height          =   495
         Left            =   360
         TabIndex        =   8
         Top             =   3000
         Width           =   1935
      End
      Begin VB.Label lblSunQuizQuestion3 
         BackColor       =   &H80000012&
         Caption         =   "The Sun produces unimaginable power through a process called__________."
         ForeColor       =   &H00FFFFFF&
         Height          =   495
         Left            =   120
         TabIndex        =   12
         Top             =   240
         Width           =   2895
      End
   End
   Begin VB.Frame fraSunQ4 
      BackColor       =   &H80000012&
      Caption         =   "Question 4"
      ForeColor       =   &H00FFFFFF&
      Height          =   3615
      Left            =   3480
      TabIndex        =   1
      Top             =   3840
      Width           =   3255
      Begin VB.OptionButton optSunQ4A1 
         BackColor       =   &H80000012&
         Caption         =   "Nitrogen and Oxygen"
         ForeColor       =   &H00FFFFFF&
         Height          =   495
         Left            =   240
         TabIndex        =   5
         Top             =   840
         Width           =   1935
      End
      Begin VB.OptionButton optSunQ4A3 
         BackColor       =   &H80000012&
         Caption         =   "Helium and Hydrogen"
         ForeColor       =   &H00FFFFFF&
         Height          =   495
         Left            =   240
         TabIndex        =   4
         Top             =   2280
         Width           =   1935
      End
      Begin VB.OptionButton optSunQ4A2 
         BackColor       =   &H80000012&
         Caption         =   "Helium and Oxygen"
         ForeColor       =   &H00FFFFFF&
         Height          =   495
         Left            =   240
         TabIndex        =   3
         Top             =   1560
         Width           =   1935
      End
      Begin VB.OptionButton optSunQ4A4 
         BackColor       =   &H80000012&
         Caption         =   "Nitrogen and  Hydrogen"
         ForeColor       =   &H00FFFFFF&
         Height          =   495
         Left            =   240
         TabIndex        =   2
         Top             =   3000
         Width           =   1935
      End
      Begin VB.Label lblSunQuizQuestion4 
         BackColor       =   &H80000012&
         Caption         =   "The Sun converts what two elements into energy?"
         ForeColor       =   &H00FFFFFF&
         Height          =   495
         Left            =   120
         TabIndex        =   6
         Top             =   240
         Width           =   3015
      End
   End
   Begin VB.CommandButton cmdSunQuizRefuel 
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
Attribute VB_Name = "frmSunQuiz"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'=======================================================================================================================================================================
'Matthew Lal
'June 5, 2010
'Sun Quiz
'The user refuels to continue the adventure by answer the questions correctly.
'=============================================================================================================================
Option Explicit

Private Sub cmdSunQuizRefuel_Click()
    'Determines what answers are correct and what answers are not.
    If optSunQ1A2.Value = True Then
    frmGalaxy3.intScoreQ1 = frmGalaxy3.intScoreQ1 + 1
        Else
        frmGalaxy3.intScoreQ1 = frmGalaxy3.intScoreQ1 - 1
    End If
    
    If optSunQ2A2.Value = True Then
    frmGalaxy3.intScoreQ2 = frmGalaxy3.intScoreQ2 + 1
        Else
        frmGalaxy3.intScoreQ2 = frmGalaxy3.intScoreQ2 - 1
    End If
    
    If optSunQ3A4.Value = True Then
    frmGalaxy3.intScoreQ3 = frmGalaxy3.intScoreQ3 + 1
        Else
        frmGalaxy3.intScoreQ3 = frmGalaxy3.intScoreQ3 - 1
    End If
    
    If optSunQ4A3.Value = True Then
    frmGalaxy3.intScoreQ4 = frmGalaxy3.intScoreQ4 + 1
        Else
        frmGalaxy3.intScoreQ4 = frmGalaxy3.intScoreQ4 - 1
    End If
    
    If optSunQ5A1.Value = True Then
    frmGalaxy3.intScoreQ5 = frmGalaxy3.intScoreQ5 + 1
        Else
        frmGalaxy3.intScoreQ5 = frmGalaxy3.intScoreQ5 - 1
    End If
    
    'Determines the users overall score for the current quiz.
    frmGalaxy3.intScore = frmGalaxy3.intScoreQ1 + frmGalaxy3.intScoreQ2 + frmGalaxy3.intScoreQ3 + frmGalaxy3.intScoreQ4 + frmGalaxy3.intScoreQ5
    
    'Adds the users score to the fuel.
    frmGalaxy1.intFuel = frmGalaxy1.intFuel + frmGalaxy3.intScore
    
    'Closes the quiz and returns to the galaxy map.
    Unload frmSunQuiz
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



