VERSION 5.00
Begin VB.Form frmGalaxy1 
   BackColor       =   &H00000000&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Galactic Adventures"
   ClientHeight    =   8445
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   10230
   Icon            =   "frmGalaxy1.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "frmGalaxy1.frx":1272
   ScaleHeight     =   8445
   ScaleWidth      =   10230
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdName 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Insert Name"
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
      Left            =   4440
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   7800
      Width           =   1455
   End
   Begin VB.CommandButton cmdNext1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Next"
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
      Left            =   7920
      MaskColor       =   &H00E0E0E0&
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   7800
      Width           =   1455
   End
   Begin VB.Label lblDialog1 
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Hello Captain, this is your first mission and you need a name for the crew to call you by. What is it?"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   975
      Left            =   840
      TabIndex        =   1
      Top             =   6600
      Width           =   8535
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Hello Captain!"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   27.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1575
      Left            =   6480
      TabIndex        =   0
      Top             =   1440
      Width           =   2295
   End
   Begin VB.Image Image1 
      BorderStyle     =   1  'Fixed Single
      Height          =   6330
      Left            =   360
      Picture         =   "frmGalaxy1.frx":6AEF
      Stretch         =   -1  'True
      Top             =   120
      Width           =   9495
   End
End
Attribute VB_Name = "frmGalaxy1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'=======================================================================================================================================================================
'Matthew Lal
'June 5, 2010
'Introduction screen.
'Introduces the user to the game and asks them for their name.
'=============================================================================================================================
Option Explicit

'All variables must be declared.
Public strPlayerName As String
Public intFuel As Integer
Dim intCount As Integer

Private Sub cmdName_Click()
intCount = intCount + 1
'Once clicked a message box appears for players name.
    strPlayerName = InputBox("What will the crew call you?", "Galactic Adventures")
End Sub

Private Sub cmdNext1_Click()
    'Makes sure the user enters a name before continuing.
    If strPlayerName = "" Then
        MsgBox ("Please insert name first!")
    ElseIf intCount >= 1 Then
    frmGalaxy1.Hide
    frmGalaxy2.Show
    End If
End Sub

Private Sub Form_Load()
    'Gives the user 100 fuel to start with.
    frmGalaxy1.intFuel = frmGalaxy1.intFuel + 100
End Sub
