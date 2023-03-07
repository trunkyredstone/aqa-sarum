'Skeleton Program code for the AQA COMP1 Summer 2015 examination
'this code would be used in conjunction with the Preliminary Material
'written by the AQA COMP1 Programmer Team
'developed in the Visual Studio 2008 programming environment

' modified by Romano Lucchesi

Module CaptureTheSarrum
    Const BoardDimension As Integer = 8

    Sub Main()
        Dim board(BoardDimension, BoardDimension) As String
        Dim gameOver As Boolean
        Dim startSquare As Integer
        Dim finishSquare As Integer
        Dim startRank As Integer
        Dim startFile As Integer
        Dim finishRank As Integer
        Dim finishFile As Integer
        Dim piecesTaken As Integer
        Dim points As Integer
        Dim moveIsLegal As Boolean
        Dim playAgain As Char
        Dim sampleGame As Char
        Dim whoseTurn As Char
        
        Do
            whoseTurn = "W"
            sampleGame = GetTypeOfGame()
            InitialiseBoard(board, sampleGame)
            Do
                DisplayBoard(board)
                DisplayWhoseTurnItIs(whoseTurn)
                Console.WriteLine($"You've taken {piecesTaken} pieces. You have {points} points.")
                Do
                    GetMove(startSquare, finishSquare)
                    startRank = startSquare Mod 10
                    startFile = startSquare\10
                    finishRank = finishSquare Mod 10
                    finishFile = finishSquare\10
                    moveIsLegal = CheckMoveIsLegal(board, startRank, startFile, finishRank, finishFile, whoseTurn)
                    If Not moveIsLegal Then
                        Console.WriteLine("That is not a legal move - please try again")
                    End If
                Loop Until moveIsLegal
                gameOver = CheckIfGameWillBeWon(board, finishRank, finishFile)
                If CheckIfPieceWillBeTaken(board, finishRank, finishFile) Then
                    piecesTaken += 1
                    points += GetPiecePoints(board, finishRank, finishFile)
                End If
                MakeMove(board, startRank, startFile, finishRank, finishFile, whoseTurn)
                If gameOver Then
                    DisplayWinner(whoseTurn)
                End If
                If whoseTurn = "W" Then
                    whoseTurn = "B"
                Else
                    whoseTurn = "W"
                End If
            Loop Until gameOver
            Console.Write("Do you want to play again (enter Y for Yes)? ")
            playAgain = Console.ReadLine
            If Asc(playAgain) >= 97 And Asc(playAgain) <= 122 Then
                playAgain = Chr(Asc(playAgain) - 32)
            End If
        Loop Until playAgain <> "Y"
    End Sub

    Private Sub DisplayWhoseTurnItIs(whoseTurn As Char)
        If whoseTurn = "W" Then
            Console.WriteLine("It is White's turn")
        Else
            Console.WriteLine("It is Black's turn")
        End If
    End Sub

    Private Function GetTypeOfGame() As Char
        Dim typeOfGame As Char
        Console.Write("Do you want to play the sample game (enter Y for Yes)? ")
        typeOfGame = Console.ReadLine
        Return typeOfGame
    End Function

    Private Sub DisplayWinner(whoseTurn As Char)
        If whoseTurn = "W" Then
            Console.WriteLine("Black's Sarrum has been captured.  White wins!")
        Else
            Console.WriteLine("White's Sarrum has been captured.  Black wins!")
        End If
        Console.WriteLine()
    End Sub

    Private Function CheckIfGameWillBeWon(board(,) As String, finishRank As Integer, finishFile As Integer) As Boolean
        If board(finishRank, finishFile)(1) = "S" Then
            Return True
        Else
            Return False
        End If
    End Function
    
    Private Function CheckIfPieceWillBeTaken(board(,) As String, finishRank As Integer, finishFile As Integer) As Boolean
        Return Not board(finishRank, finishFile)(1) = " "
    End Function
    
    Private Function GetPiecePoints(board(,) As String, finishRank As Integer, finishFile As Integer) As Integer
        Dim piece = board(finishRank, finishFile)(1)
        
        If piece = "S" Then
            Return 1
        ElseIf piece = "N" Then
            Return 2
        ElseIf piece = "E" Then
            Return 3
        ElseIf piece = "G" Then
            Return 4
        ElseIf piece = "R" Then
            Return 5
        ElseIf piece = "K" Then
            Return 6
        End If
        
        Return 0
    End Function

    Private Sub DisplayBoard(board(,) As String)
        Dim rankNo As Integer
        Dim fileNo As Integer
        Console.WriteLine()
        For rankNo = 1 To BoardDimension
            Console.WriteLine("    _________________________")
            Console.Write(rankNo & "   ")
            For fileNo = 1 To BoardDimension
                Console.Write("|" & board(rankNo, fileNo))
            Next
            Console.WriteLine("|")
        Next
        Console.WriteLine("    _________________________")
        Console.WriteLine()
        Console.WriteLine("     1  2  3  4  5  6  7  8")
        Console.WriteLine()
        Console.WriteLine()
    End Sub

    Private Function CheckRedumMoveIsLegal(board(,) As String, startRank As Integer, startFile As Integer,
                                           finishRank As Integer,
                                           finishFile As Integer, colourOfPiece As Char) As Boolean
        If colourOfPiece = "W" Then
            If finishRank = startRank - 1 Then
                If finishFile = startFile And board(finishRank, finishFile) = "  " Then
                    Return True
                ElseIf Math.Abs(finishFile - startFile) = 1 And board(finishRank, finishFile)(0) = "B" Then
                    Return True
                End If
            End If
        Else
            If finishRank = startRank + 1 Then
                If finishFile = startFile And board(finishRank, finishFile) = "  " Then
                    Return True
                ElseIf Math.Abs(finishFile - startFile) = 1 And board(finishRank, finishFile)(0) = "W" Then
                    Return True
                End If
            End If
        End If
        Return False
    End Function

    Private Function CheckSarrumMoveIsLegal(startRank As Integer, startFile As Integer,
                                            finishRank As Integer, finishFile As Integer) As Boolean
        If Math.Abs(finishFile - startFile) <= 1 And Math.Abs(finishRank - startRank) <= 1 Then
            Return True
        End If
        Return False
    End Function

    Private Function CheckGisgigirMoveIsLegal(board(,) As String, startRank As Integer, startFile As Integer,
                                              finishRank As Integer, finishFile As Integer) As Boolean
        Dim gisgigirMoveIsLegal As Boolean
        Dim count As Integer
        Dim rankDifference As Integer
        Dim fileDifference As Integer
        gisgigirMoveIsLegal = False
        rankDifference = finishRank - startRank
        fileDifference = finishFile - startFile
        If rankDifference = 0 Then
            If fileDifference >= 1 Then
                gisgigirMoveIsLegal = True
                For count = 1 To fileDifference - 1
                    If board(startRank, startFile + count) <> "  " Then
                        gisgigirMoveIsLegal = False
                    End If
                Next
            ElseIf fileDifference <= - 1 Then
                gisgigirMoveIsLegal = True
                For count = - 1 To fileDifference + 1 Step - 1
                    If board(startRank, startFile + count) <> "  " Then
                        gisgigirMoveIsLegal = False
                    End If
                Next
            End If
        ElseIf fileDifference = 0 Then
            If rankDifference >= 1 Then
                gisgigirMoveIsLegal = True
                For count = 1 To rankDifference - 1
                    If board(startRank + count, startFile) <> "  " Then
                        gisgigirMoveIsLegal = False
                    End If
                Next
            ElseIf rankDifference <= - 1 Then
                gisgigirMoveIsLegal = True
                For count = - 1 To rankDifference + 1 Step - 1
                    If board(startRank + count, startFile) <> "  " Then
                        gisgigirMoveIsLegal = False
                    End If
                Next
            End If
        End If
        Return gisgigirMoveIsLegal
    End Function

    Private Function CheckNabuMoveIsLegal(startRank As Integer, startFile As Integer, finishRank As Integer,
                                          finishFile As Integer) As Boolean
        If Math.Abs(finishFile - startFile) <= 2 And Math.Abs(finishRank - startRank) <= 2 Then
            Return True
        End If
        Return False
    End Function

    Private Function CheckEtluMoveIsLegal(startRank As Integer, startFile As Integer,
                                          finishRank As Integer,
                                          finishFile As Integer) As Boolean
        If _
            Math.Abs(finishFile - startFile) = 2 And Math.Abs(finishRank - startRank) = 1 Or
            Math.Abs(finishFile - startFile) = 1 And Math.Abs(finishRank - startRank) = 2 Then
            Return True
        End If
        Return False
    End Function
    
    Private Function CheckKashaptuMoveIsLegal(board(,) As String, startRank As Integer, startFile As Integer,
                                              finishRank As Integer, finishFile As Integer, colourOfPiece As Char) As Boolean
        ' Check move is horizontal, vertical, or diagonal
        If Not (Math.Abs(finishFile - startFile) = Math.Abs(finishRank - startRank) Or Math.Abs(finishFile - startFile) = 0 Or Math.Abs(finishRank - startRank) = 0) Then
            Return False
        End If
        
        ' Check for obstructions
        If Math.Abs(finishFile - startFile) = Math.Abs(finishRank - startRank) Then ' Diagonal
            Dim rankDifference = 1
            Dim currentRank = startRank
            
            If startRank > finishRank Then
                rankDifference = -1
            End If
            
            If finishFile > startFile Then ' Top Right
                For currentFile = startFile to finishFile
                    If Not board(currentRank, currentFile) = "  " And Not board(currentRank, currentFile) = colourOfPiece & "K" Then Return False
                    currentRank += rankDifference
                Next
            ElseIf finishFile < startFile Then ' Top Left
                For currentFile = finishFile To startFile Step -1
                    If Not board(currentRank, currentFile) = "  " And Not board(currentRank, currentFile) = colourOfPiece & "K" Then Return False
                    currentRank += rankDifference
                Next
            End If
        ElseIf Math.Abs(finishFile - startFile) = 0 Then ' Vertical
            If finishFile > startFile Then
                For i = startFile To finishFile 
                    If Not board(startRank, i) = "  " And Not board(startRank, i)(1) = "K" Then Return False
                Next
            Else 
                For i = finishFile To startFile Step -1
                    If Not board(startRank, i) = "  " And Not board(startRank, i)(1) = "K" Then Return False
                Next
            End If
        ElseIf Math.Abs(finishRank - startRank) = 0 Then ' Horizontal
            If finishRank > startRank Then
                For i = startRank To finishRank
                    If Not board(i, startFile) = "  " And Not board(startRank, i)(1) = "K" Then Return False
                Next
            Else 
                For i = finishRank To startRank Step -1
                    If Not board(i, startFile) = "  " And Not board(startRank, i)(1) = "K" Then Return False
                Next
            End If
        End If
        
        Return True
    End Function

    Private Function CheckMoveIsLegal(board(,) As String, startRank As Integer, startFile As Integer,
                                      finishRank As Integer,
                                      finishFile As Integer, whoseTurn As Char) As Boolean
        Dim pieceType As Char
        Dim pieceColour As Char
        If finishFile = startFile And finishRank = startRank Then
            Return False
        End If
        pieceType = board(startRank, startFile)(1)
        pieceColour = board(startRank, startFile)(0)
        If whoseTurn = "W" Then
            If pieceColour <> "W" Then
                Return False
            End If
            If board(finishRank, finishFile)(0) = "W" Then
                Return False
            End If
        Else
            If pieceColour <> "B" Then
                Return False
            End If
            If board(finishRank, finishFile)(0) = "B" Then
                Return False
            End If
        End If
        Select Case pieceType
            Case "R"
                Return CheckRedumMoveIsLegal(board, startRank, startFile, finishRank, finishFile, pieceColour)
            Case "S"
                Return CheckSarrumMoveIsLegal(startRank, startFile, finishRank, finishFile)
            Case "G"
                Return CheckGisgigirMoveIsLegal(board, startRank, startFile, finishRank, finishFile)
            Case "N"
                Return CheckNabuMoveIsLegal(startRank, startFile, finishRank, finishFile)
            Case "E"
                Return CheckEtluMoveIsLegal(startRank, startFile, finishRank, finishFile)
            Case "K"
                Return CheckKashaptuMoveIsLegal(board, startRank, startFile, finishRank, finishFile, pieceColour)
            Case Else
                Return False
        End Select
    End Function

    Private Sub InitialiseBoard(ByRef board(,) As String, sampleGame As Char)
        Dim rankNo As Integer
        Dim fileNo As Integer
        If sampleGame = "Y" Then
            For rankNo = 1 To BoardDimension
                For fileNo = 1 To BoardDimension
                    board(rankNo, fileNo) = "  "
                Next
            Next
            board(1, 2) = "BG"
            board(1, 4) = "BS"
            board(1, 8) = "WG"
            board(2, 1) = "WR"
            board(3, 1) = "WS"
            board(3, 2) = "BE"
            board(3, 8) = "BE"
            board(6, 8) = "BR"
            board(1, 1) = "BK"
            board(2, 4) = "WK"
        Else
            For rankNo = 1 To BoardDimension
                For fileNo = 1 To BoardDimension
                    If rankNo = 2 Then
                        board(rankNo, fileNo) = "BR"
                    ElseIf rankNo = 7 Then
                        board(rankNo, fileNo) = "WR"
                    ElseIf rankNo = 1 Or rankNo = 8 Then
                        If rankNo = 1 Then board(rankNo, fileNo) = "B"
                        If rankNo = 8 Then board(rankNo, fileNo) = "W"
                        Select Case fileNo
                            Case 1, 8
                                board(rankNo, fileNo) = board(rankNo, fileNo) & "G"
                            Case 2, 7
                                board(rankNo, fileNo) = board(rankNo, fileNo) & "E"
                            Case 3, 6
                                board(rankNo, fileNo) = board(rankNo, fileNo) & "N"
                            Case 4
                                board(rankNo, fileNo) = board(rankNo, fileNo) & "K"
                            Case 5
                                board(rankNo, fileNo) = board(rankNo, fileNo) & "S"
                        End Select
                    Else
                        board(rankNo, fileNo) = "  "
                    End If
                Next
            Next
        End If
    End Sub

    Private Sub GetMove(ByRef startSquare As Integer, ByRef finishSquare As Integer)
        Console.Write("Enter coordinates of square containing piece to move (file first): ")
        startSquare = Console.ReadLine
        Console.Write("Enter coordinates of square to move piece to (file first): ")
        finishSquare = Console.ReadLine
    End Sub

    Private Sub MakeMove(ByRef board(,) As String, startRank As Integer, startFile As Integer, finishRank As Integer,
                         finishFile As Integer, whoseTurn As Char)
        If whoseTurn = "W" And finishRank = 1 And board(startRank, startFile)(1) = "R" Then
            board(finishRank, finishFile) = "W" & ChoosePiece()
            board(startRank, startFile) = "  "
        ElseIf whoseTurn = "B" And finishRank = 8 And board(startRank, startFile)(1) = "R" Then
            board(finishRank, finishFile) = "B" & ChoosePiece()
            board(startRank, startFile) = "  "
        Else
            board(finishRank, finishFile) = board(startRank, startFile)
            board(startRank, startFile) = "  "
        End If
    End Sub
    
    Private Function ChoosePiece() As Char
        Dim piece As Char
        Dim choiceValid = False
        
        Do Until choiceValid
            Console.Write("Choose which piece to convert to: ")
            piece = Char.ToUpper(Console.ReadLine)
            
            If piece = "S" Or piece = "N" Or piece = "E" Or piece = "G" Or piece = "R" Or piece = "K" Then
                choiceValid = True
            End If
        Loop
        
        Return piece
    End Function
End Module
