VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "Msflxgrd.ocx"
Begin VB.Form frmGame 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Minesweeper"
   ClientHeight    =   9810
   ClientLeft      =   1995
   ClientTop       =   6045
   ClientWidth     =   12285
   Icon            =   "MinesweeperV2_JasonZhen.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   9810
   ScaleWidth      =   12285
   Begin VB.Timer tmrTime 
      Interval        =   1000
      Left            =   360
      Top             =   3360
   End
   Begin MSFlexGridLib.MSFlexGrid grdGame 
      Height          =   1815
      Left            =   960
      TabIndex        =   0
      Top             =   1200
      Width           =   2175
      _ExtentX        =   3836
      _ExtentY        =   3201
      _Version        =   393216
      Rows            =   9
      Cols            =   9
      FixedRows       =   0
      FixedCols       =   0
      FocusRect       =   0
      ScrollBars      =   0
      BorderStyle     =   0
   End
   Begin VB.Image Image1 
      Height          =   240
      Left            =   1080
      Picture         =   "MinesweeperV2_JasonZhen.frx":08CA
      Top             =   8280
      Visible         =   0   'False
      Width           =   240
   End
   Begin VB.Label lblBombsLeft 
      Alignment       =   2  'Center
      BorderStyle     =   1  'Fixed Single
      Caption         =   "00"
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   375
      Left            =   360
      TabIndex        =   2
      Top             =   600
      Width           =   735
   End
   Begin VB.Label lblTime 
      Alignment       =   2  'Center
      BorderStyle     =   1  'Fixed Single
      Caption         =   "000"
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   375
      Left            =   2520
      TabIndex        =   1
      Top             =   600
      Width           =   735
   End
   Begin VB.Image imgFace 
      Height          =   390
      Left            =   1320
      Picture         =   "MinesweeperV2_JasonZhen.frx":0CCF
      Top             =   480
      Width           =   390
   End
   Begin VB.Image imgFlag 
      Height          =   255
      Left            =   9600
      Picture         =   "MinesweeperV2_JasonZhen.frx":12D3
      Stretch         =   -1  'True
      Top             =   6240
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.Image imgBomb 
      Height          =   255
      Index           =   2
      Left            =   9600
      Picture         =   "MinesweeperV2_JasonZhen.frx":16FD
      Stretch         =   -1  'True
      Top             =   5760
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.Image imgBomb 
      Height          =   255
      Index           =   1
      Left            =   9600
      Picture         =   "MinesweeperV2_JasonZhen.frx":1B34
      Stretch         =   -1  'True
      Top             =   5280
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.Image imgBomb 
      Height          =   255
      Index           =   0
      Left            =   9600
      Picture         =   "MinesweeperV2_JasonZhen.frx":1F3D
      Stretch         =   -1  'True
      Top             =   4800
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.Image imgUntouched 
      Height          =   255
      Left            =   9600
      Picture         =   "MinesweeperV2_JasonZhen.frx":2336
      Stretch         =   -1  'True
      Top             =   7200
      Visible         =   0   'False
      Width           =   240
   End
   Begin VB.Image imgQuestion 
      Height          =   255
      Left            =   9600
      Picture         =   "MinesweeperV2_JasonZhen.frx":26D0
      Stretch         =   -1  'True
      Top             =   6720
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.Image imgSmiley 
      Height          =   390
      Index           =   4
      Left            =   6960
      Picture         =   "MinesweeperV2_JasonZhen.frx":2AD4
      Top             =   5400
      Visible         =   0   'False
      Width           =   390
   End
   Begin VB.Image imgSmiley 
      Height          =   375
      Index           =   3
      Left            =   6480
      Picture         =   "MinesweeperV2_JasonZhen.frx":30D5
      Top             =   6360
      Visible         =   0   'False
      Width           =   390
   End
   Begin VB.Image imgSmiley 
      Height          =   390
      Index           =   2
      Left            =   6480
      Picture         =   "MinesweeperV2_JasonZhen.frx":36CD
      Top             =   5880
      Visible         =   0   'False
      Width           =   390
   End
   Begin VB.Image imgSmiley 
      Height          =   375
      Index           =   1
      Left            =   6480
      Picture         =   "MinesweeperV2_JasonZhen.frx":3CE4
      Top             =   6840
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.Image imgSmiley 
      Height          =   390
      Index           =   0
      Left            =   6480
      Picture         =   "MinesweeperV2_JasonZhen.frx":4297
      Top             =   5400
      Visible         =   0   'False
      Width           =   390
   End
   Begin VB.Image imgSprite 
      Height          =   255
      Index           =   8
      Left            =   9120
      Picture         =   "MinesweeperV2_JasonZhen.frx":489B
      Stretch         =   -1  'True
      Top             =   8640
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.Image imgSprite 
      Height          =   255
      Index           =   7
      Left            =   9120
      Picture         =   "MinesweeperV2_JasonZhen.frx":4C45
      Stretch         =   -1  'True
      Top             =   8160
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.Image imgSprite 
      Height          =   255
      Index           =   6
      Left            =   9120
      Picture         =   "MinesweeperV2_JasonZhen.frx":502F
      Stretch         =   -1  'True
      Top             =   7680
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.Image imgSprite 
      Height          =   255
      Index           =   5
      Left            =   9120
      Picture         =   "MinesweeperV2_JasonZhen.frx":5422
      Stretch         =   -1  'True
      Top             =   7200
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.Image imgSprite 
      Height          =   255
      Index           =   4
      Left            =   9120
      Picture         =   "MinesweeperV2_JasonZhen.frx":581C
      Stretch         =   -1  'True
      Top             =   6720
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.Image imgSprite 
      Height          =   255
      Index           =   3
      Left            =   9120
      Picture         =   "MinesweeperV2_JasonZhen.frx":5C23
      Stretch         =   -1  'True
      Top             =   6240
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.Image imgSprite 
      Height          =   255
      Index           =   2
      Left            =   9120
      Picture         =   "MinesweeperV2_JasonZhen.frx":6025
      Stretch         =   -1  'True
      Top             =   5760
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.Image imgSprite 
      Height          =   255
      Index           =   1
      Left            =   9120
      Picture         =   "MinesweeperV2_JasonZhen.frx":6423
      Stretch         =   -1  'True
      Top             =   5280
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.Image imgSprite 
      Height          =   255
      Index           =   0
      Left            =   9120
      Picture         =   "MinesweeperV2_JasonZhen.frx":682A
      Stretch         =   -1  'True
      Top             =   4800
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.Menu mnuMenu 
      Caption         =   "Menu"
      Begin VB.Menu mnuAbout 
         Caption         =   "About"
      End
      Begin VB.Menu mnuBar2 
         Caption         =   "-"
      End
      Begin VB.Menu mnuNew 
         Caption         =   "New Game"
      End
      Begin VB.Menu mnuHighScores 
         Caption         =   "HighScores"
      End
      Begin VB.Menu mnuBar1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuBeginner 
         Caption         =   "Beginner"
         Checked         =   -1  'True
      End
      Begin VB.Menu mnuIntermediate 
         Caption         =   "Intermediate "
      End
      Begin VB.Menu mnuExpert 
         Caption         =   "Expert"
      End
   End
End
Attribute VB_Name = "frmGame"
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

Dim Space() As Integer
Dim Rows As Integer
Dim Cols As Integer
Dim NumBombs As Integer
Dim Hard(1 To MAX_SCORES) As Integer
Dim Med(1 To MAX_SCORES) As Integer
Dim Beg(1 To MAX_SCORES) As Integer
Dim Level As Integer

Private Sub Form_Load()
   
    Rows = 9
    Cols = 9
    NumBombs = 10
    Level = 1
    
    'Create an array that represents each space on the grid.
    
    ReDim Space(0 To Rows - 1, 0 To Cols - 1) As Integer
    
    tmrTime.Enabled = False
    lblBombsLeft = NumBombs

    InitGame grdGame, Rows, Cols, imgUntouched, frmGame, imgFace, Space(), lblTime, lblBombsLeft
    GenBombs Space(), NumBombs, Rows, Cols
    GetScores Hard(), Med(), Beg()
    
End Sub

Private Sub grdGame_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
    If Button = 1 Then
        imgFace.Picture = imgSmiley(3)
    End If
End Sub

Private Sub grdGame_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
    Dim Row As Integer
    Dim Col As Integer
    Dim Score As Integer
    
    Col = grdGame.MouseCol
    Row = grdGame.MouseRow
    
    grdGame.Col = Col
    grdGame.Row = Row
    
    imgFace.Picture = imgSmiley(0)
    
    tmrTime.Enabled = True
    
    'Make the board invisible, load the images, and then make it visible again to
    'make the images appear instantly.
    
    grdGame.Visible = False
    
    If Button = 1 And grdGame.CellPicture <> imgFlag Then
        
        If Space(Row, Col) >= UNEXPLORED Then
            If Space(Row, Col) = UNEXPLORED Then
                ExpandSpaces Space, Row, Col, grdGame, imgUntouched
                DisplayGrid Space(), grdGame, imgSprite(), imgUntouched
            
            ElseIf Space(Row, Col) > UNEXPLORED Then
                Set grdGame.CellPicture = imgSprite(Space(Row, Col))
                
            End If
            
            If CheckWin(grdGame, NumBombs, imgUntouched, imgFlag, imgQuestion) Then
                grdGame.Visible = True
                grdGame.Enabled = False
            
                imgFace.Picture = imgSmiley(2)
                tmrTime.Enabled = False
                
                Score = Int(lblTime.Caption)
                
                If Level = 1 Then
                    CheckScore Score, Beg(), MAX_SCORES
                ElseIf Level = 2 Then
                    CheckScore Score, Med(), MAX_SCORES
                ElseIf Level = 3 Then
                    CheckScore Score, Hard(), MAX_SCORES
                End If
                
                SaveSCores Hard(), Med(), Beg(), MAX_SCORES
                
            End If
            
        ElseIf Space(Row, Col) = Bomb Then
        
            Set grdGame.CellPicture = imgBomb(1)
            GameOver Space(), imgBomb(), grdGame, imgFlag, Row, Col
            
            imgFace.Picture = imgSmiley(4)
            tmrTime.Enabled = False
            
        End If
        
    ElseIf Button = 2 Then

        If Space(Row, Col) >= -1 And grdGame.CellPicture = imgUntouched Then
            Set grdGame.CellPicture = imgFlag
            
            lblBombsLeft.Caption = Int(lblBombsLeft.Caption) - 1
            
        ElseIf grdGame.CellPicture = imgFlag Then
            Set grdGame.CellPicture = imgQuestion
            
            lblBombsLeft.Caption = Int(lblBombsLeft.Caption) + 1
        ElseIf grdGame.CellPicture = imgQuestion Then
            Set grdGame.CellPicture = imgUntouched
        End If
    End If
    
    grdGame.Visible = True
End Sub

Private Sub imgFace_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
    If Button = 1 Then
        imgFace.Picture = imgSmiley(1)
    End If
End Sub

Private Sub imgFace_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
    If Button = 1 Then
        imgFace.Picture = imgSmiley(0)
        
        InitGame grdGame, Rows, Cols, imgUntouched, frmGame, imgFace, Space(), lblTime, lblBombsLeft
        GenBombs Space(), NumBombs, Rows, Cols
        
        grdGame.Enabled = True
        
        tmrTime.Enabled = False
        lblTime.Caption = "000"
        
        lblBombsLeft = NumBombs
    End If
End Sub

Private Sub mnuAbout_Click()
    frmAbout.Show vbModal
End Sub

Private Sub mnuBeginner_Click()
    mnuBeginner.Checked = True
    mnuIntermediate.Checked = False
    mnuExpert.Checked = False
    
    Rows = 9
    Cols = 9
    NumBombs = 10

    ReDim Space(0 To Rows - 1, 0 To Cols - 1) As Integer
    
    imgFace.Picture = imgSmiley(0)
    
    InitGame grdGame, Rows, Cols, imgUntouched, frmGame, imgFace, Space(), lblTime, lblBombsLeft
    GenBombs Space(), NumBombs, Rows, Cols
    
    grdGame.Enabled = True
    
    tmrTime.Enabled = False
    lblTime.Caption = "000"
    
    lblBombsLeft = NumBombs
    Level = 1
End Sub

Private Sub mnuExpert_Click()
    mnuBeginner.Checked = False
    mnuIntermediate.Checked = False
    mnuExpert.Checked = True
    
    Rows = 30
    Cols = 16
    NumBombs = 99
    
    ReDim Space(0 To Rows - 1, 0 To Cols - 1) As Integer
    
    imgFace.Picture = imgSmiley(0)
    
    InitGame grdGame, Rows, Cols, imgUntouched, frmGame, imgFace, Space(), lblTime, lblBombsLeft
    GenBombs Space(), NumBombs, Rows, Cols
    
    grdGame.Enabled = True
    
    tmrTime.Enabled = False
    lblTime.Caption = "000"
    
    lblBombsLeft = NumBombs
    Level = 3
End Sub

Private Sub mnuHighScores_Click()
    frmHighScores.Show vbModal
End Sub

Private Sub mnuIntermediate_Click()
    mnuBeginner.Checked = False
    mnuIntermediate.Checked = True
    mnuExpert.Checked = False
    
    Rows = 16
    Cols = 16
    NumBombs = 40
    
    ReDim Space(0 To Rows - 1, 0 To Cols - 1) As Integer
    
    imgFace.Picture = imgSmiley(0)
    
    InitGame grdGame, Rows, Cols, imgUntouched, frmGame, imgFace, Space(), lblTime, lblBombsLeft
    GenBombs Space(), NumBombs, Rows, Cols
    
    grdGame.Enabled = True
    
    tmrTime.Enabled = False
    lblTime.Caption = "000"
    
    lblBombsLeft = NumBombs
    Level = 2
End Sub

Private Sub mnuNew_Click()
    
    ReDim Space(0 To Rows - 1, 0 To Cols - 1) As Integer
    
    imgFace.Picture = imgSmiley(0)
    
    InitGame grdGame, Rows, Cols, imgUntouched, frmGame, imgFace, Space(), lblTime, lblBombsLeft
    GenBombs Space(), NumBombs, Rows, Cols
    
    grdGame.Enabled = True
    
    tmrTime.Enabled = False
    lblTime.Caption = "000"
    
    lblBombsLeft = NumBombs
End Sub


'Private Sub tmrTime_Timer()
'    Dim Time As Integer
'
'    Time = lblTime.Caption
'    Time = Time + 1
'    lblTime.Caption = VBA.Format$(Time, "000")
'
'End Sub
