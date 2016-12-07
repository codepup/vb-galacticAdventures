VERSION 5.00
Begin VB.Form frmShipInfo 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Galactic Adventures - Ship Information"
   ClientHeight    =   3030
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   4560
   Icon            =   "frmShipInfo.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "frmShipInfo.frx":1272
   ScaleHeight     =   3030
   ScaleWidth      =   4560
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdReturn 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Close"
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
      Left            =   240
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   2400
      Width           =   1335
   End
   Begin VB.CommandButton cmdExit 
      BackColor       =   &H00FFFFFF&
      Caption         =   "E&xit Game"
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
      Left            =   2880
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   2400
      Width           =   1335
   End
   Begin VB.Label lblHint 
      Alignment       =   2  'Center
      BackColor       =   &H00000000&
      Caption         =   "Hint: Before landing, hover over planets."
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   480
      TabIndex        =   3
      Top             =   120
      Width           =   3495
   End
   Begin VB.Image imgShipInfo 
      Height          =   1140
      Left            =   1080
      Picture         =   "frmShipInfo.frx":6AEF
      Top             =   480
      Width           =   2250
   End
   Begin VB.Label lblFuelDisplay 
      Alignment       =   2  'Center
      BackColor       =   &H00000000&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   375
      Left            =   1440
      TabIndex        =   2
      Top             =   1920
      Width           =   1575
   End
End
Attribute VB_Name = "frmShipInfo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'=======================================================================================================================================================================
'Matthew Lal
'June 5, 2010
'Ship Information
'Displays information on the users ship, like fuel.
'=============================================================================================================================
Option Explicit
Dim intDummy As Integer

Private Sub cmdExit_Click()
    'Allows the user to exit the game from here.
    intDummy = MsgBox("Are You Sure?", vbYesNo)
    
    If intDummy = vbYes Then
        End
        ElseIf intDummy = vbNo Then
7
    End If
End Sub

Private Sub cmdReturn_Click()
    'Closes the ship inof screen and returns to the galaxy map.
    Unload frmShipInfo
    frmGalaxy3.Show
End Sub

Private Sub Form_Load()
    'Displays the users current fuel our of 100.
    lblFuelDisplay.Caption = "Fuel = " & frmGalaxy1.intFuel & " / 100"
End Sub
