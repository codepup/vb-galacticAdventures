VERSION 5.00
Begin VB.Form frmWarning3 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Galactic Adventures"
   ClientHeight    =   8445
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   10080
   Icon            =   "frmWarning3.frx":0000
   LinkTopic       =   "Form3"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "frmWarning3.frx":1272
   ScaleHeight     =   8445
   ScaleWidth      =   10080
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdNext2 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Roger"
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
      TabIndex        =   0
      Top             =   7800
      Width           =   1455
   End
   Begin VB.Label lblDialog2 
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Radar is giving us a clear reading ... something else is out there ... watching us; to the west. Keep it slow Captain."
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Left            =   840
      TabIndex        =   2
      Top             =   6600
      Width           =   8535
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Radar is back..."
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   27.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1455
      Left            =   6600
      TabIndex        =   1
      Top             =   1560
      Width           =   2055
   End
   Begin VB.Image Image1 
      BorderStyle     =   1  'Fixed Single
      Height          =   6330
      Left            =   360
      Picture         =   "frmWarning3.frx":6AEF
      Stretch         =   -1  'True
      Top             =   120
      Width           =   9495
   End
End
Attribute VB_Name = "frmWarning3"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'=======================================================================================================================================================================
'Matthew Lal
'June 5, 2010
'Warning
'Warns the player of the secret of the galaxy.
'=============================================================================================================================
Option Explicit

Private Sub cmdNext2_Click()
    'Closes the warning and returns to the galaxy map.
    Unload frmWarning3
    frmGalaxy3.Show
End Sub


