Attribute VB_Name = "mainmod"
Public Const Player_Human As Byte = 0
Public Const Player_CPUEASY As Byte = 1
Public Const Player_CPUNORM As Byte = 2
Public Const Player_CPUHARD As Byte = 3
Public Const Player_CPUEXPT As Byte = 4
Public Const Player_CPUBETA As Byte = 5
Public Const Player_TWOCPUS As Byte = 6

Public Type BoardType
    GameIn As Boolean
    movesleft As Integer
    Turn As Byte
    Player(1) As Byte
    Value(100) As Integer
    Allowed(100) As Boolean
    FrameDir As Integer
    FramePos As Integer
    Score(1) As Integer
    moves(100) As Boolean
    Last As Integer
    maxplys(1) As Integer
    eval As Integer
    maxonrow As Integer
End Type
Public GameIn As Boolean
Public Current_Board As BoardType
Public gameclosed As Boolean

Public Type DiffLevelType
    Name As String
    BasePlys As Integer
    ExtPlys As Integer
    ELO As Integer
    Singular_Extension As Boolean
End Type
Public dbg As Boolean
Public DiffLevel(5) As DiffLevelType
Public DebugInfo As String
Public ClockCycles As Long
Public Nodes As Long

Public ScoreBoard(4) As New board

Public Sub DrawBoard(Tboard As BoardType)
For a = 0 To 99
    mainform.box(a).Caption = Tboard.Value(a)
    If Tboard.Allowed(a) = True Then
        mainform.box(a).BackStyle = 1
        mainform.box(a).ForeColor = RGB(255, 255, 255)
        If mainform.box(a).Caption > 9 Then
            mainform.box(a).BackColor = RGB(80, 194, 80)
        Else
            grad = (mainform.box(a).Caption + 9) / 18
            If grad = 0 Then grad = 0.001
            If grad = 1 Then grad = 0.999
            bc = 1 - grad
            gc = grad
            r = Int((80 * gc) + (255 * bc))
            g = Int((80 * gc) + (0 * bc))
            b = Int((194 * gc) + (0 * bc))
            mainform.box(a).BackColor = RGB(r, g, b)
        End If
        mainform.box(a).Visible = True
    Else
        mainform.box(a).Caption = ""
        mainform.box(a).BackStyle = 0
        mainform.box(a).Visible = True
    End If
Next a
If Tboard.Last <> 100 Then mainform.box(Tboard.Last).ForeColor = RGB(200, 200, 200)
If Tboard.Turn = 0 Then
    mainform.turncur.Caption = "<-"
    If Tboard.Player(Tboard.Turn) = Player_Human Then
        mainform.stat.Caption = "Player One's turn."
    Else
        mainform.stat.Caption = "Computers turn... (thinking!)"
    End If
Else
    mainform.turncur.Caption = "->"
    If Tboard.Player(Tboard.Turn) = Player_Human Then
        mainform.stat.Caption = "Player Two's turn."
    Else
        mainform.stat.Caption = "Computers turn... (thinking!)"
    End If
End If
mainform.frame.Visible = False
If Tboard.FrameDir = 0 Then
    mainform.frame.Width = 4935
    mainform.frame.Height = 375
    mainform.frame.Top = mainform.box(10 * Tboard.FramePos).Top
    mainform.frame.Left = 120
Else
    mainform.frame.Width = 375
    mainform.frame.Height = 4935
    mainform.frame.Top = 120
    mainform.frame.Left = mainform.box(Tboard.FramePos).Left
End If
mainform.frame.Visible = True
mainform.sco(0) = Tboard.Score(0)
mainform.sco(1) = Tboard.Score(1)
If Tboard.GameIn = False Then
    MsgBox "Player 1: " & Tboard.Score(0) & vbCrLf & "Player 2: " & Tboard.Score(1), , "Game Over !!!"
    If Tboard.Player(1) > Player_Human And Tboard.Player(1) <= Player_CPUEXPT Then
        If Tboard.Score(0) < 1 Then
            Tboard.Score(1) = Tboard.Score(1) + (1 - Tboard.Score(0))
            Tboard.Score(0) = 1
        ElseIf Tboard.Score(1) < 1 Then
            Tboard.Score(0) = Tboard.Score(0) + (1 - Tboard.Score(1))
            Tboard.Score(1) = 1
        End If
        If Tboard.Score(0) < Tboard.Score(1) Then
            Tboard.Score(0) = Tboard.Score(0) - 55
            Tboard.Score(1) = Tboard.Score(1) + 55
            Tboard.Score(1) = Tboard.Score(1) + (10 - Tboard.Score(0))
            Tboard.Score(0) = Tboard.Score(0) + (10 - Tboard.Score(0))
        ElseIf Tboard.Score(0) > Tboard.Score(1) Then
            Tboard.Score(0) = Tboard.Score(0) + 50
            Tboard.Score(1) = Tboard.Score(1) - 50
            Tboard.Score(0) = Tboard.Score(0) + (10 - Tboard.Score(1))
            Tboard.Score(1) = Tboard.Score(1) + (10 - Tboard.Score(1))
        End If
        pscore = Round(((Tboard.Score(0) / (Tboard.Score(0) + Tboard.Score(1))) * 110) - 55, 0)
        If Tboard.Player(0) = Player_Human Then
            pname$ = USmooth(InputBox("Your score was: " & pscore & vbCrLf & "What is your name?", "Player Name Entry", "NoName"))
        ElseIf Tboard.Player(0) = Player_CPUEASY Then
            pname$ = USmooth(DiffLevel(1).Name)
        ElseIf Tboard.Player(0) = Player_CPUNORM Then
            pname$ = USmooth(DiffLevel(2).Name)
        ElseIf Tboard.Player(0) = Player_CPUHARD Then
            pname$ = USmooth(DiffLevel(3).Name)
        ElseIf Tboard.Player(0) = Player_CPUEXPT Then
            pname$ = USmooth(DiffLevel(4).Name)
        ElseIf Tboard.Player(0) = Player_CPUBETA Then
            pname$ = USmooth(DiffLevel(5).Name)
        End If
        If pname$ <> "" Then
            ScoreBoard(Tboard.Player(1) - 1).AddEntry pname$, pscore
            SaveScores
            MsgBox pname$ & "           " & pscore & vbCrLf & "Vs" & vbCrLf & _
            USmooth(DiffLevel(Tboard.Player(1)).Name) & "           " & (0) - pscore
            'mainform.menu_highscore_show_Click
        End If
    End If
    Exit Sub
End If
If Tboard.Player(Tboard.Turn) > Player_Human Then mainform.cputurn.Enabled = True Else mainform.cputurn.Enabled = False
End Sub

Public Sub SetBoard(Tboard As BoardType, ByVal Level As Byte)
DebugInfo = ""
DiffLevel(0).BasePlys = 2
DiffLevel(0).ExtPlys = 0
DiffLevel(0).ELO = 1500
DiffLevel(0).Singular_Extension = False
DiffLevel(1).BasePlys = 2
DiffLevel(1).ExtPlys = 0
DiffLevel(1).ELO = 1192
DiffLevel(1).Singular_Extension = False
DiffLevel(2).BasePlys = 2
DiffLevel(2).ExtPlys = 2
DiffLevel(2).ELO = 1517
DiffLevel(2).Singular_Extension = False
DiffLevel(3).BasePlys = 2
DiffLevel(3).ExtPlys = 4
DiffLevel(3).ELO = 1792
DiffLevel(3).Singular_Extension = False
Select Case ClockCycles
Case Is < 15000: DiffLevel(4).BasePlys = 0: DiffLevel(4).ExtPlys = 2
Case Is < 90000: DiffLevel(4).BasePlys = 0: DiffLevel(4).ExtPlys = 4
Case Is < 8100000: DiffLevel(4).BasePlys = 0: DiffLevel(4).ExtPlys = 6
Case Else: DiffLevel(4).BasePlys = 0: DiffLevel(4).ExtPlys = 8
End Select
DiffLevel(4).Singular_Extension = False
Select Case ClockCycles
Case Is < 15000: DiffLevel(5).BasePlys = 0: DiffLevel(5).ExtPlys = 2
Case Is < 90000: DiffLevel(5).BasePlys = 0: DiffLevel(5).ExtPlys = 4
Case Is < 8100000: DiffLevel(5).BasePlys = 0: DiffLevel(5).ExtPlys = 6
Case Else: DiffLevel(5).BasePlys = 0: DiffLevel(5).ExtPlys = 8
End Select
DiffLevel(5).Singular_Extension = True
Tboard.GameIn = True
Tboard.movesleft = 100
Tboard.Turn = Int(Rnd(1) * 2)
Tboard.Score(0) = 0
Tboard.Score(1) = 0
Tboard.Player(0) = Player_Human
Tboard.Player(1) = Level
If Level = Player_TWOCPUS Then
    Do
        Tboard.Player(0) = Int(Rnd(1) * 5) + 1
        Tboard.Player(1) = Int(Rnd(1) * 4) + 1
    Loop Until Tboard.Player(0) <> Tboard.Player(1)
    mainform.Caption = "Rulne - " & DiffLevel(Tboard.Player(0)).Name & " Vs " & DiffLevel(Tboard.Player(1)).Name
End If
For a = 0 To 1
    Tboard.maxplys(a) = 0 - DiffLevel(Tboard.Player(a)).ExtPlys
Next a
Do
    DoEvents
    total = 0
    For a = 0 To 99
        Tboard.Allowed(a) = True
        Tboard.Value(a) = Int(Rnd(1) * 101)
        Select Case Tboard.Value(a)
        Case Is < 1: Tboard.Value(a) = -9
        Case Is < 3: Tboard.Value(a) = -8
        Case Is < 6: Tboard.Value(a) = -7
        Case Is < 10: Tboard.Value(a) = -6
        Case Is < 15: Tboard.Value(a) = -5
        Case Is < 21: Tboard.Value(a) = -4
        Case Is < 28: Tboard.Value(a) = -3
        Case Is < 36: Tboard.Value(a) = -2
        Case Is < 45: Tboard.Value(a) = -1
        Case Is < 55: Tboard.Value(a) = 1
        Case Is < 65: Tboard.Value(a) = 2
        Case Is < 76: Tboard.Value(a) = 3
        Case Is < 80: Tboard.Value(a) = 4
        Case Is < 86: Tboard.Value(a) = 5
        Case Is < 91: Tboard.Value(a) = 6
        Case Is < 95: Tboard.Value(a) = 7
        Case Is < 98: Tboard.Value(a) = 8
        Case Is < 101: Tboard.Value(a) = 9
        End Select
        If Tboard.Value(a) = 0 Then Tboard.Value(a) = 1
        total = total + Tboard.Value(a)
        'bsize = 1
        'Select Case Int(a / 10)
        'Case Is < bsize, Is > 9 - bsize: Tboard.Allowed(a) = False: Tboard.Value(a) = 0
        'End Select
        'Select Case a - (Int(a / 10)) * 10
        'Case Is < bsize, Is > 9 - bsize: Tboard.Allowed(a) = False: Tboard.Value(a) = 0
        'End Select
    Next a
    Tboard.Last = 100
    Tboard.FrameDir = Int(Rnd(1) * 2)
    Tboard.FramePos = Int(Rnd(1) * 10)
Loop Until total > 90 And total < 110
Tboard.Value(Int(Rnd(1) * 100)) = 10
Tboard.Value(Int(Rnd(1) * 100)) = 15
movnum = 0
If Tboard.FrameDir = 0 Then ' check frame for possible moves
    curoff = Tboard.FramePos * 10
    For a = 0 To 9 ' for entire row
        Tboard.moves(curoff + a) = True
        If Tboard.Allowed(curoff + a) = True Then movnum = movnum + 1
    Next a
Else
    For a = 0 To 9 ' for entire column
        Tboard.moves((a * 10) + Tboard.FramePos) = True
        If Tboard.Allowed((a * 10) + Tboard.FramePos) = True Then movnum = movnum + 1
    Next a
End If
Tboard.Last = 100
If movnum = 0 Then
    Tboard.GameIn = False
End If
End Sub

Public Sub DoMove(Tboard As BoardType, ByVal Square As Integer)
If Tboard.Allowed(Square) = False Then Exit Sub
If Tboard.FrameDir = 0 Then ' check chosen box is in frame
    squarepos = Int(Square / 10)
    If squarepos <> Tboard.FramePos Then Exit Sub
Else
    squarepos = Square - (Int(Square / 10) * 10)
    If squarepos <> Tboard.FramePos Then Exit Sub
End If
Tboard.movesleft = Tboard.movesleft - 1
Tboard.Score(Tboard.Turn) = Tboard.Score(Tboard.Turn) + Tboard.Value(Square) ' add/subtract score to player score
Tboard.Allowed(Square) = False ' disable chosen box
If Tboard.FrameDir = 0 Then ' rotate frame
    Tboard.FrameDir = 1
    Tboard.FramePos = Square - (Int(Square / 10) * 10)
Else
    Tboard.FrameDir = 0
    Tboard.FramePos = Int(Square / 10)
End If
If Tboard.Turn = 1 Then Tboard.Turn = 0 Else Tboard.Turn = 1 ' swap turns
For a = 0 To 99 ' blank-out all allowed moves
    Tboard.moves(a) = False
Next a
movnum = 0
Tboard.maxonrow = -10
If Tboard.FrameDir = 0 Then ' check frame for possible moves
    curoff = Tboard.FramePos * 10
    For a = 0 To 9 ' for entire row
        Tboard.moves(curoff + a) = True
        If Tboard.Allowed(curoff + a) = True Then
            movnum = movnum + 1
            If Tboard.Value(curoff + a) > Tboard.maxonrow Then Tboard.maxonrow = Tboard.Value(curoff + a)
        End If
    Next a
Else
    For a = 0 To 9 ' for entire column
        Tboard.moves((a * 10) + Tboard.FramePos) = True
        If Tboard.Allowed((a * 10) + Tboard.FramePos) = True Then
            movnum = movnum + 1
            If Tboard.Value((a * 10) + Tboard.FramePos) > Tboard.maxonrow Then Tboard.maxonrow = Tboard.Value((a * 10) + Tboard.FramePos)
        End If
    Next a
End If
Tboard.Last = Square
If movnum = 0 Then
    Tboard.GameIn = False
End If
End Sub

Public Function DoCompMove(Tboard As BoardType, Score As Integer, ByVal plysleft As Integer) As Integer
hnum = 0
hval = -16000
plysleft = plysleft - 2
Dim ttboard As BoardType
Dim moves, scoleft As Integer
Dim tttboard As BoardType
Dim cureval As Integer
scoleft = 1
moves = 0
    For a = 0 To 99
        If Tboard.moves(a) = True Then scoleft = scoleft + 1
        If Tboard.moves(a) = True And Tboard.Allowed(a) = True Then moves = moves + 1
    Next a
If DiffLevel(Tboard.Player(Tboard.Turn)).Singular_Extension Then
    If moves = 1 And plysleft = 0 Then plysleft = plysleft + 1
End If
For a = 0 To 99
    If Tboard.moves(a) = True And Tboard.Allowed(a) = True Then
        Nodes = Nodes + 1
        DoEvents
        ttboard = Tboard
        DoMove ttboard, a
        ttlow = 16000
        cureval = (ttboard.Score(0) - ttboard.Score(1)) * 3
        If Tboard.Turn = 0 Then
            cureval = cureval + ttboard.Turn
        Else
            cureval = cureval - ttboard.Turn
        End If
        If Tboard.Turn = 1 Then cureval = 0 - cureval
        If ttboard.GameIn = False Then
            ttlow = cureval * 100
        Else
            scoleft = 1
            moves = 0
                For b = 0 To 99
                    If ttboard.moves(b) = True Then scoleft = scoleft + 1
                    If ttboard.moves(b) = True And ttboard.Allowed(b) = True Then moves = moves + 1
                Next b
            If DiffLevel(Tboard.Player(Tboard.Turn)).Singular_Extension Then
                If moves = 1 And plysleft = 0 Then plysleft = plysleft + 1
            End If
            For b = 0 To 99
                If ttboard.moves(b) = True And ttboard.Allowed(b) = True Then
                    Nodes = Nodes + 1
                    tttboard = ttboard
                    DoMove tttboard, b
                    cureval = (tttboard.Score(0) - tttboard.Score(1)) * 3
                    If Tboard.Turn = 0 Then
                        cureval = cureval + tttboard.Turn
                    Else
                        cureval = cureval - tttboard.Turn
                    End If
                    If Tboard.Turn = 1 Then cureval = 0 - cureval
                    If tttboard.GameIn = False Then
                        cureval = cureval * 100
                    ElseIf plysleft > 0 Or (plysleft > Tboard.maxplys(Tboard.Turn) And cureval <= ttlow And cureval >= hval) Then
                        c = DoCompMove(tttboard, cureval, plysleft)
                    End If
                    If cureval <= ttlow Then ttlow = cureval
                    
                End If
            Next b
        End If
        If plysleft = DiffLevel(Tboard.Player(Tboard.Turn)).BasePlys - 2 Then
            DebugInfo = DebugInfo & Tboard.Value(a) & " (" & ttlow & ")" & vbCrLf
        End If
        If ttlow >= hval Then
            hval = ttlow: hnum = a
        End If
        If hval > 300 Then Exit For
    End If
    If Tboard.GameIn = False Then Exit Function
    If gameclosed = True Then End
    Next a
If plysleft = DiffLevel(Tboard.Player(Tboard.Turn)).BasePlys - 2 Then
    DebugInfo = DebugInfo & "Best: " & Tboard.Value(hnum) & " (" & hval & ")" & vbCrLf
End If
Tboard.eval = hval
DoCompMove = hnum
Score = hval
End Function

Public Sub condense(ByVal g As Integer)
Dim Score As New board
Dim Newscore As New board
Dim Done(31) As String
Dim numnames As Integer
Dim nxt As Boolean
Score.HighScoreFile = "C:\WINDOWS\rulne" & g & ".dat"
Score.LoadHighScoreData
Newscore.NewHighScoreData False
Newscore.BestLow = False
If Score.NumberOfEntrys = 0 Then Exit Sub
numnames = 0
For a = 1 To 30
    Done(a) = ""
Next a
Do
    b = b + 1
    curname = Score.Name(b)
    nxt = False
    For a = 1 To 30
        If LSmooth(curname) = LSmooth(Done(a)) Then nxt = True: Exit For
    Next a
    If nxt = False Then
        numnames = numnames + 1
        Done(numnames) = curname
        Newscore.AddEntry curname, Score.Score(b)
    End If
Loop Until b = Score.NumberOfEntrys
Newscore.HighScoreFile = Score.HighScoreFile
Newscore.SaveHighScoreData
End Sub

Public Sub LoadScores()
For a = 0 To 3
    ScoreBoard(a).HighScoreFile = "c:\windows\rulne" & a & ".dat"
    If FileExist(ScoreBoard(a).HighScoreFile) = True Then
        ScoreBoard(a).LoadHighScoreData
    Else
        ScoreBoard(a).NewHighScoreData False
        ScoreBoard(a).SaveHighScoreData
    End If
Next a
End Sub

Public Sub SaveScores()
For a = 0 To 3
    ScoreBoard(a).HighScoreFile = "c:\windows\rulne" & a & ".dat"
    ScoreBoard(a).SaveHighScoreData
Next a
End Sub
