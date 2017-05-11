VERSION 5.00
Begin VB.Form frmAbout 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "About"
   ClientHeight    =   2160
   ClientLeft      =   14115
   ClientTop       =   6375
   ClientWidth     =   3030
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2160
   ScaleWidth      =   3030
   Begin VB.CommandButton cmdOk 
      Caption         =   "Ok"
      Height          =   495
      Left            =   1560
      TabIndex        =   1
      Top             =   1560
      Width           =   1215
   End
   Begin VB.Image imgSmiley 
      Height          =   495
      Index           =   0
      Left            =   120
      Picture         =   "frmAbout.frx":0000
      Stretch         =   -1  'True
      Top             =   120
      Width           =   495
   End
   Begin VB.Label Label2 
      Caption         =   "Copyright © 2015-2016 Jason Zhen"
      Height          =   375
      Left            =   120
      TabIndex        =   2
      Top             =   1200
      Width           =   2895
   End
   Begin VB.Label Label1 
      Caption         =   "Minesweeper For Windows"
      Height          =   255
      Left            =   120
      TabIndex        =   0
      Top             =   840
      Width           =   2295
   End
End
Attribute VB_Name = "frmAbout"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Title: Minesweeper V2
'Author: Jason Zhen
'Date: June 2, 2016
'Files: MinesweeperV2_JasonZhen.vbp, Minesweeper(Splash)V2_JasonZhen.frm, Minesweeper(Splash)V2_JasonZhen.frx,
'       MinesweeperV2_JasonZhen.frm , MinesweeperV2_JasonZhen.frx, MinesweeperV2_JasonZhen.bas, HighScores.txt
'       frmHighScores.frm, frmAbout.frm, frmAbout.frx
'Purpose: The purpose of this program is to play the game Minesweeper.

Option Explicit

Private Sub cmdOk_Click()
    frmAbout.Hide
    Unload Me
End Sub

Private Sub Form_Load()
    With frmAbout
        .Top = (Screen.Height - .Height) / 2
        .Left = (Screen.Width - .Width) / 2
    End With
End Sub

