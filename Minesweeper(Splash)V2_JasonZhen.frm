VERSION 5.00
Begin VB.Form frmSplash 
   BorderStyle     =   0  'None
   Caption         =   "Minesweeper"
   ClientHeight    =   4740
   ClientLeft      =   2205
   ClientTop       =   4920
   ClientWidth     =   10560
   Icon            =   "Minesweeper(Splash)V2_JasonZhen.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   4740
   ScaleWidth      =   10560
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Timer Timer1 
      Interval        =   500
      Left            =   8760
      Top             =   240
   End
   Begin VB.Label Label1 
      Caption         =   "MINESWEEPER Ver. 1.0  Copyright 2016 Jason Zhen"
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   36
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3735
      Left            =   3720
      TabIndex        =   0
      Top             =   480
      Width           =   6375
   End
   Begin VB.Image Image1 
      Height          =   2580
      Left            =   480
      Picture         =   "Minesweeper(Splash)V2_JasonZhen.frx":08CA
      Stretch         =   -1  'True
      Top             =   480
      Width           =   3180
   End
End
Attribute VB_Name = "frmSplash"
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

Private Sub Form_Load()
    With frmSplash
        .Top = (Screen.Height - .Height) / 2
        .Left = (Screen.Width - .Width) / 2
    End With
    Timer1.Enabled = True
End Sub

Private Sub Timer1_Timer()
    frmGame.Show
    Timer1.Enabled = False
    Unload frmSplash
End Sub

