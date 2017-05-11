Attribute VB_Name = "Module1"
'Title: Minesweeper V2
'Author: Jason Zhen
'Date: June 2, 2016
'Files: MinesweeperV2_JasonZhen.vbp, Minesweeper(Splash)V2_JasonZhen.frm, Minesweeper(Splash)V2_JasonZhen.frx,
'       MinesweeperV2_JasonZhen.frm , MinesweeperV2_JasonZhen.frx, MinesweeperV2_JasonZhen.bas
'Purpose: The purpose of this program is to play the game Minesweeper.

Option Explicit

Global Const MAX_SCORES = 5
Global Const Bomb = -1
Global Const EXPLORED = -2
Global Const UNEXPLORED = 0

Public Sub GetScores(Hard() As Integer, Med() As Integer, Beg() As Integer)

    Dim K As Integer
    
    K = 0
    Open App.Path & "\Highscores.txt" For Input As #1
        Do While Not (EOF(1))
            
            K = K + 1
            
            Input #1, Hard(K)
            Input #1, Med(K)
            Input #1, Beg(K)
        Loop
    Close #1
End Sub

Public Sub SaveSCores(Hard() As Integer, Med() As Integer, Beg() As Integer, Max As Integer)
    
    Dim K As Integer
    
    Open App.Path & "\Highscores.txt" For Output As #1
        For K = 1 To Max
            Write #1, Hard(K)
            Write #1, Med(K)
            Write #1, Beg(K)
        Next K
    Close #1
    
End Sub

'This procedure randomly generates bombs in unique places on the grid.

Public Sub GenBombs(Space() As Integer, NumBombs As Integer, Rows As Integer, Cols As Integer)
    
    Dim K As Integer
    Dim BRow As Integer
    Dim BCol As Integer
    Dim Valid As Boolean
    
    For K = 1 To NumBombs
        
        Valid = False
        Do
            BRow = RandomInt(Rows - 1, 0)
            BCol = RandomInt(Cols - 1, 0)
            If Space(BRow, BCol) >= UNEXPLORED Then
                Valid = True
                Space(BRow, BCol) = Bomb
                
                'Call the procedure to increase all the value of all adjacent spaces to the newly
                'created bomb by one.
                
                IncrementNeighbours Space(), BRow, BCol, Rows, Cols
            End If
        Loop While Not Valid
    Next K
    
End Sub

'This procedure increases the value of each neighbour space. The value of a space corresponds
'to how many bombs are adjacent to it.

Public Sub IncrementNeighbours(Space() As Integer, ByVal Row As Integer, _
                               ByVal Col As Integer, ByVal Rows As Integer, _
                               ByVal Cols As Integer)
                               
    Dim R As Integer
    Dim C As Integer
    Dim Row1 As Integer
    Dim Row2 As Integer
    Dim Col1 As Integer
    Dim Col2 As Integer
    
    'Alter the range of the loop depending on where the player clicks.
    'For example if the user clicks on an edge then there will only be five adjacent spaces instead
    'of eight.
    
    If Row = 0 Then
        Row1 = Row
        Row2 = Row + 1
    ElseIf Row = Rows - 1 Then
        Row1 = Row - 1
        Row2 = Row
    Else
        Row1 = Row - 1
        Row2 = Row + 1
    End If
    
    If Col = 0 Then
        Col1 = Col
        Col2 = Col + 1
    ElseIf Col = Cols - 1 Then
        Col1 = Col - 1
        Col2 = Col
    Else
        Col1 = Col - 1
        Col2 = Col + 1
    End If
    
    
    For R = Row1 To Row2
        For C = Col1 To Col2
            If Space(R, C) >= UNEXPLORED Then
                Space(R, C) = Space(R, C) + 1
            End If
        Next C
    Next R
        
End Sub

Public Function RandomInt(High As Integer, Low As Integer) As Integer

    RandomInt = Int(Rnd * (High - Low + 1) + Low)
End Function

'This procedure will initialize the values of the space array, the size and picture of each cell,
'the position of the grid, the position of the timer and bombs left labels, and the smiley
'image.

Public Sub InitGame(Grid As Control, Rows As Integer, Cols As Integer, Image As Image, _
                    Form As Form, Face As Image, Space() As Integer, Time As Label, _
                    BombsLeft As Label)
                    
    Dim R As Integer
    Dim C As Integer
    Dim ExtraWidth As Integer
    Dim ExtraHeight As Integer
    
    Grid.Visible = False
    
    ExtraWidth = 15
    ExtraHeight = 15
    
    Grid.Rows = Rows
    Grid.Cols = Cols
    
    For R = 0 To Rows - 1
    
        Grid.RowHeight(R) = Image.Height
        Grid.Row = R
        For C = 0 To Cols - 1
            Space(R, C) = UNEXPLORED
            Grid.Col = C
            Grid.ColWidth(C) = Image.Width
            Set Grid.CellPicture = Image
        Next C
    Next R
    
    Grid.Width = Cols * (Grid.CellWidth + ExtraWidth)
    Grid.Height = Rows * (Grid.CellHeight + ExtraHeight)
    
    Form.Height = Grid.Height + 2000
    Form.Width = Grid.Width + 1750
    
    Grid.Left = Form.ScaleWidth / 2 - Grid.Width / 2
    Grid.Top = Form.ScaleHeight / 2 - Grid.Height / 2
    
    Face.Left = Grid.Left + Grid.Width / 2 - Face.Width / 2
    Face.Top = Grid.Top - Face.Height * 1.5
    
    Time.Left = Grid.Left + Grid.Width - Time.Width
    Time.Top = Grid.Top - Time.Height
    
    BombsLeft.Left = Grid.Left
    BombsLeft.Top = Grid.Top - BombsLeft.Height
    
    Grid.Visible = True
    
    With Form
        .Top = (Screen.Height - .Height) / 2
        .Left = (Screen.Width - .Width) / 2
    End With
End Sub

'This procedure will display images that represent the values of each cell.

Public Sub DisplayGrid(Space() As Integer, Grid As Control, Sprite As Variant, Untouched As Image)
    
    Dim R As Integer
    Dim C As Integer
    Dim Rows As Integer
    Dim Cols As Integer
    
    Rows = Grid.Rows
    Cols = Grid.Cols
    
    For R = 0 To Rows - 1
        For C = 0 To Cols - 1
            Grid.Col = C
            Grid.Row = R
            
            If Space(R, C) = EXPLORED And Grid.CellPicture = Untouched Then
                Set Grid.CellPicture = Sprite(0)
                ShowNeighbours Space(), R, C, Sprite, Grid, Untouched
            End If
        Next C
    Next R
End Sub

'This procedure reveals the value of each cell.

Public Sub ShowNeighbours(Space() As Integer, ByVal Row As Integer, ByVal Col As Integer, _
                          Sprite As Variant, Grid As Control, Untouched As Image)
                          
    Dim R As Integer
    Dim C As Integer
    Dim Row1 As Integer
    Dim Row2 As Integer
    Dim Col1 As Integer
    Dim Col2 As Integer
    
    If Row = 0 Then
        Row1 = Row
        Row2 = Row + 1
    ElseIf Row = Grid.Rows - 1 Then
        Row1 = Row - 1
        Row2 = Row
    Else
        Row1 = Row - 1
        Row2 = Row + 1
    End If
    
    If Col = 0 Then
        Col1 = Col
        Col2 = Col + 1
    ElseIf Col = Grid.Cols - 1 Then
        Col1 = Col - 1
        Col2 = Col
    Else
        Col1 = Col - 1
        Col2 = Col + 1
    End If
    
    
    For R = Row1 To Row2
        Grid.Row = R
        For C = Col1 To Col2
            Grid.Col = C
            If Space(R, C) > UNEXPLORED And Grid.CellPicture = Untouched Then
                
                Set Grid.CellPicture = Sprite(Space(R, C))
            End If
        Next C
    Next R
    
    
End Sub

'This procedure recursively opens adjacent spaces until it reaches the edge of the grid, or a non-empty cell.

Public Sub ExpandSpaces(Space() As Integer, ByVal Row As Integer, ByVal Col As Integer, _
                        Grid As Control, Untouched As Image)
    
    Dim R As Integer
    Dim C As Integer
    Dim Row1 As Integer
    Dim Row2 As Integer
    Dim Col1 As Integer
    Dim Col2 As Integer
    
    If Row = 0 Then
        Row1 = Row
        Row2 = Row + 1
    ElseIf Row = Grid.Rows - 1 Then
        Row1 = Row - 1
        Row2 = Row
    Else
        Row1 = Row - 1
        Row2 = Row + 1
    End If
    
    If Col = 0 Then
        Col1 = Col
        Col2 = Col + 1
    ElseIf Col = Grid.Cols - 1 Then
        Col1 = Col - 1
        Col2 = Col
    Else
        Col1 = Col - 1
        Col2 = Col + 1
    End If
    
    
    For R = Row1 To Row2
        Grid.Row = R
        For C = Col1 To Col2
            Grid.Col = C
            If Space(R, C) = UNEXPLORED And Grid.CellPicture = Untouched Then
                Space(R, C) = EXPLORED
                ExpandSpaces Space(), R, C, Grid, Untouched
            End If
        Next C
    Next R
                    
End Sub

'This procedure will reveal all bombs when the user loses and show incorrect flag placements.

Public Sub GameOver(Space() As Integer, ImageBomb As Variant, Grid As Control, _
                    Flag As Image, BRow As Integer, BCol As Integer)

    Dim R As Integer
    Dim C As Integer
    
    For R = 0 To Grid.Rows - 1
        Grid.Row = R
        For C = 0 To Grid.Cols - 1
            Grid.Col = C
            If Space(R, C) = Bomb And (R <> BRow Or C <> BCol) And Grid.CellPicture <> Flag Then
                Set Grid.CellPicture = ImageBomb(0)
                
            ElseIf Space(R, C) <> Bomb And Grid.CellPicture = Flag Then
                Set Grid.CellPicture = ImageBomb(2)
            End If
        Next C
    Next R
    
    Grid.Enabled = False
End Sub

'This procedure will check if the user has beaten the game. The game is won when all spaces, except
'for spaces where bombs exist, are uncovered.

Public Function CheckWin(Grid As Control, NumBombs As Integer, UnOpened As Image, _
                         Flag As Image, Question As Image) As Boolean

    Dim C As Integer
    Dim R As Integer
    Dim Count As Integer
    
    Count = 0
    CheckWin = False
    
    For C = 0 To Grid.Cols - 1
        Grid.Col = C
        
        For R = 0 To Grid.Rows - 1
            Grid.Row = R
             
            If Grid.CellPicture = UnOpened Or Grid.CellPicture = Flag _
            Or Grid.CellPicture = Question Then
                Count = Count + 1
            End If
            
        Next R
    Next C
    
    If Count = NumBombs Then
    
        For C = 0 To Grid.Cols - 1
            Grid.Col = C
        
            For R = 0 To Grid.Rows - 1
                Grid.Row = R
                 
                If Grid.CellPicture = UnOpened Or Grid.CellPicture = Question Then
                    Grid.CellPicture = Flag
                End If
            
            Next R
        Next C
        CheckWin = True
    End If
    
End Function

Public Sub CheckScore(Score As Integer, HighScore() As Integer, Max As Integer)
    Dim K As Integer
    Dim Finished As Boolean
    
    K = 1
    Do
        Finished = False
        If Score > HighScore(K) Then
            HighScore(K) = Score
            Finished = True
        Else
            K = K + 1
        End If
        
    Loop Until Finished Or K = Max + 1
End Sub




















