VERSION 5.00
Begin VB.Form frmHome 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Galactic Adventures"
   ClientHeight    =   4470
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6000
   Icon            =   "frmHome.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "frmHome.frx":1272
   ScaleHeight     =   4470
   ScaleWidth      =   6000
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdExit 
      BackColor       =   &H00FFFFFF&
      Caption         =   "E&xit"
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
      Left            =   4560
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   3480
      Width           =   1095
   End
   Begin VB.CommandButton cmdHelp 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Help"
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
      Left            =   3120
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   3480
      Width           =   1095
   End
   Begin VB.CommandButton cmdAbout 
      BackColor       =   &H00FFFFFF&
      Caption         =   "About"
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
      Left            =   1680
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   3480
      Width           =   1095
   End
   Begin VB.CommandButton cmdPlay 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Play Now!"
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
      MaskColor       =   &H00000000&
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   3480
      Width           =   1095
   End
   Begin VB.Image Image1 
      Height          =   615
      Left            =   240
      Picture         =   "frmHome.frx":2DD2
      Top             =   360
      Width           =   5685
   End
End
Attribute VB_Name = "frmHome"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub cmdAbout_Click()
    frmHome.Hide
    frmAbout.Show
End Sub

Private Sub cmdExit_Click()
    End
End Sub

Private Sub cmdHelp_Click()
    frmHome.Hide
    frmHelp.Show
End Sub

Private Sub cmdPlay_Click()
    frmHome.Hide
    frmGalaxy1.Show
End Sub
