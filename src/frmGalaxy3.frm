VERSION 5.00
Begin VB.Form frmGalaxy3 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Galactic Adventures"
   ClientHeight    =   8955
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   11970
   Icon            =   "frmGalaxy3.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   MouseIcon       =   "frmGalaxy3.frx":1272
   MousePointer    =   99  'Custom
   Picture         =   "frmGalaxy3.frx":1B3C
   ScaleHeight     =   8955
   ScaleWidth      =   11970
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdShipInfo 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Ship Information"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   9840
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   8280
      Width           =   1935
   End
   Begin VB.Image imgAliens 
      Height          =   1215
      Left            =   8040
      ToolTipText     =   "ALERT! Our radar is getting ridiculous readings!"
      Top             =   7080
      Width           =   1335
   End
   Begin VB.Image imgClue3 
      Height          =   3135
      Left            =   9720
      ToolTipText     =   "ALERT! Something is out there!"
      Top             =   4920
      Width           =   1695
   End
   Begin VB.Image imgClue2 
      Height          =   1935
      Left            =   7920
      ToolTipText     =   "ALERT! Radar is getting fuzzy and showing false readings!"
      Top             =   2520
      Width           =   2655
   End
   Begin VB.Image imgClue1 
      Height          =   2055
      Left            =   5760
      ToolTipText     =   "ALERT! Radar is showing false readings!"
      Top             =   2400
      Width           =   2055
   End
   Begin VB.Image imgSun 
      Height          =   3135
      Left            =   0
      ToolTipText     =   "Gather Intel and resources on this star."
      Top             =   0
      Width           =   3855
   End
   Begin VB.Image imgJupiter 
      Height          =   2415
      Left            =   240
      ToolTipText     =   "Gather Intel and resources on this planet."
      Top             =   6360
      Width           =   2895
   End
   Begin VB.Image imgNeptune 
      Height          =   1695
      Left            =   9360
      ToolTipText     =   "Gather Intel and resources on this planet."
      Top             =   720
      Width           =   1575
   End
   Begin VB.Image imgMars 
      Height          =   975
      Left            =   6240
      ToolTipText     =   "Gather Intel and resources on this planet. ALERT! Radar is acting weird, it's coming from the south."
      Top             =   1440
      Width           =   975
   End
   Begin VB.Image imgEarth 
      Height          =   975
      Left            =   4080
      ToolTipText     =   "Gather Intel and resources on this planet."
      Top             =   4080
      Width           =   1095
   End
   Begin VB.Image imgMercury 
      Height          =   615
      Left            =   2280
      ToolTipText     =   "Gather Intel and resources on this planet."
      Top             =   3120
      Width           =   735
   End
   Begin VB.Image imgVenus 
      Height          =   735
      Left            =   1080
      ToolTipText     =   "Gather Intel and resources on this planet."
      Top             =   4080
      Width           =   975
   End
   Begin VB.Image imgPluto 
      Height          =   855
      Left            =   11280
      ToolTipText     =   "Gather Intel and resources on this planet."
      Top             =   3960
      Width           =   615
   End
   Begin VB.Image imgSaturn 
      Height          =   2055
      Left            =   4080
      ToolTipText     =   "Gather Intel and resources on this planet."
      Top             =   5880
      Width           =   3135
   End
   Begin VB.Image imgUranus 
      Height          =   2055
      Left            =   7680
      ToolTipText     =   "Gather Intel and resources on this planet."
      Top             =   4560
      Width           =   1815
   End
End
Attribute VB_Name = "frmGalaxy3"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'=======================================================================================================================================================================
'Matthew Lal
'June 5, 2010
'Galaxy Map
'Displays all the planets that lead to quizzes to refuel the users ship.
'=======================================================================================================================================================================

Option Explicit

'All variables must be declared.
Public intScore As Integer
Public intScoreQ1 As Integer
Public intScoreQ2 As Integer
Public intScoreQ3 As Integer
Public intScoreQ4 As Integer
Public intScoreQ5 As Integer
Public intScoreQ6 As Integer
Dim intRandomNumber As Integer
Dim intDummy As Integer

Private Sub cmdShipInfo_Click()
    'Displays user ship information.
    frmShipInfo.Show
End Sub

Private Sub Form_Load()
    
    'This code creates a 2% that during your space travel the alien encounter may occur.
    Randomize
    intRandomNumber = Int(Rnd * (100 - 1) + 1)
    
    'Game over if fuel runs out.
    If frmGalaxy1.intFuel <= 0 Then
        intDummy = MsgBox("You Have Run Of Fuel! Please start the game again.", vbOKOnly)
    End If
    If intDummy = 1 Then
        End
    End If
    
End Sub

Private Sub imgAliens_Click()
    'Lowers the user's ship fuel bay by 10 gallons; at the same time, close the galaxy map and opens what the user selected.
    frmGalaxy1.intFuel = frmGalaxy1.intFuel - 10
    Unload frmGalaxy3
    frmWarning4.Show
    
End Sub


Private Sub imgClue1_Click()
    'Lowers the user's ship fuel bay by 10 gallons; at the same time, close the galaxy map and opens what the user selected.
    frmGalaxy1.intFuel = frmGalaxy1.intFuel - 10
    Unload frmGalaxy3
    frmWarning1.Show
End Sub

Private Sub imgClue2_Click()
    'Lowers the user's ship fuel bay by 10 gallons; at the same time, close the galaxy map and opens what the user selected.
    frmGalaxy1.intFuel = frmGalaxy1.intFuel - 10
    Unload frmGalaxy3
    frmWarning2.Show
End Sub

Private Sub imgClue3_Click()
    'Lowers the user's ship fuel bay by 10 gallons; at the same time, close the galaxy map and opens what the user selected.
    frmGalaxy1.intFuel = frmGalaxy1.intFuel - 10
    Unload frmGalaxy3
    frmWarning3.Show
End Sub

Private Sub imgEarth_Click()
    'Lowers the user's ship fuel bay by 10 gallons; at the same time, close the galaxy map and opens what the user selected.
    'Also may be a 10% that the user bumps into the Aliens instead of the users desired selection.
    frmGalaxy1.intFuel = frmGalaxy1.intFuel - 10
     If intRandomNumber >= 90 And intRandomNumber <= 100 Then
        frmAliens.Show
        Unload frmGalaxy3
        Else
            Unload frmGalaxy3
            frmEarth.Show
    End If
End Sub

Private Sub imgJupiter_Click()
    'Lowers the user's ship fuel bay by 10 gallons; at the same time, close the galaxy map and opens what the user selected.
    'Also may be a 10% that the user bumps into the Aliens instead of the users desired selection.
    frmGalaxy1.intFuel = frmGalaxy1.intFuel - 10
    If intRandomNumber >= 90 And intRandomNumber <= 100 Then
        frmAliens.Show
        Unload frmGalaxy3
        Else
            Unload frmGalaxy3
            frmJupiter.Show
    End If
End Sub

Private Sub imgMars_Click()
    'Lowers the user's ship fuel bay by 10 gallons; at the same time, close the galaxy map and opens what the user selected.
    'Also may be a 10% that the user bumps into the Aliens instead of the users desired selection.
    frmGalaxy1.intFuel = frmGalaxy1.intFuel - 10
    If intRandomNumber >= 90 And intRandomNumber <= 100 Then
        frmAliens.Show
        Unload frmGalaxy3
        Else
            Unload frmGalaxy3
            frmMars.Show
    End If
End Sub

Private Sub imgMercury_Click()
    'Lowers the user's ship fuel bay by 10 gallons; at the same time, close the galaxy map and opens what the user selected.
    'Also may be a 10% that the user bumps into the Aliens instead of the users desired selection.
    frmGalaxy1.intFuel = frmGalaxy1.intFuel - 10
    If intRandomNumber >= 90 And intRandomNumber <= 100 Then
        frmAliens.Show
        Unload frmGalaxy3
        Else
            Unload frmGalaxy3
            frmMercury.Show
    End If
End Sub

Private Sub imgNeptune_Click()
    'Lowers the user's ship fuel bay by 10 gallons; at the same time, close the galaxy map and opens what the user selected.
    'Also may be a 10% that the user bumps into the Aliens instead of the users desired selection.
    frmGalaxy1.intFuel = frmGalaxy1.intFuel - 10
    If intRandomNumber >= 90 And intRandomNumber <= 100 Then
        frmAliens.Show
        Unload frmGalaxy3
        Else
            Unload frmGalaxy3
            frmNeptune.Show
    End If
End Sub

Private Sub imgPluto_Click()
    'Lowers the user's ship fuel bay by 10 gallons; at the same time, close the galaxy map and opens what the user selected.
    'Also may be a 10% that the user bumps into the Aliens instead of the users desired selection.
    frmGalaxy1.intFuel = frmGalaxy1.intFuel - 10
    If intRandomNumber >= 90 And intRandomNumber <= 100 Then
        frmAliens.Show
        Unload frmGalaxy3
        Else
            Unload frmGalaxy3
            frmPluto.Show
    End If
End Sub

Private Sub imgSaturn_Click()
    'Lowers the user's ship fuel bay by 10 gallons; at the same time, close the galaxy map and opens what the user selected.
    'Also may be a 10% that the user bumps into the Aliens instead of the users desired selection.
    frmGalaxy1.intFuel = frmGalaxy1.intFuel - 10
    If intRandomNumber >= 90 And intRandomNumber <= 100 Then
        frmAliens.Show
        Unload frmGalaxy3
        Else
            Unload frmGalaxy3
            frmSaturn.Show
    End If
End Sub

Private Sub imgSun_Click()
    'Lowers the user's ship fuel bay by 10 gallons; at the same time, close the galaxy map and opens what the user selected.
    'Also may be a 10% that the user bumps into the Aliens instead of the users desired selection.
    frmGalaxy1.intFuel = frmGalaxy1.intFuel - 10
    If intRandomNumber >= 90 And intRandomNumber <= 100 Then
        frmAliens.Show
        Unload frmGalaxy3
        Else
            Unload frmGalaxy3
            frmSun.Show
    End If
End Sub

Private Sub imgUranus_Click()
    'Lowers the user's ship fuel bay by 10 gallons; at the same time, close the galaxy map and opens what the user selected.
    'Also may be a 10% that the user bumps into the Aliens instead of the users desired selection.
    frmGalaxy1.intFuel = frmGalaxy1.intFuel - 10
    If intRandomNumber >= 90 And intRandomNumber <= 100 Then
        frmAliens.Show
        Unload frmGalaxy3
        Else
            Unload frmGalaxy3
            frmUranus.Show
    End If
End Sub

Private Sub imgVenus_Click()
    'Lowers the user's ship fuel bay by 10 gallons; at the same time, close the galaxy map and opens what the user selected.
    'Also may be a 10% that the user bumps into the Aliens instead of the users desired selection.
    frmGalaxy1.intFuel = frmGalaxy1.intFuel - 10
    If intRandomNumber >= 90 And intRandomNumber <= 100 Then
        frmAliens.Show
        Unload frmGalaxy3
        Else
            Unload frmGalaxy3
            frmVenus.Show
    End If
End Sub
