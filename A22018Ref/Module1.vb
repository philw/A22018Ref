' Skeleton Program code for the AQA A Level Paper 1 2018 examination
' this code should be used in conjunction with the Preliminary Material
' written by the AQA Programmer Team
' developed using Visual Studio Community Edition

Class QueueOfTiles
    Protected Contents() As Char
    Protected Rear As Integer
    Protected MaxSize As Integer

    Public Sub New(ByVal MaxSize As Integer)
        Randomize()
        Rear = -1
        Me.MaxSize = MaxSize
        ReDim Contents(Me.MaxSize - 1)
        For Count = 0 To Me.MaxSize - 1
            Contents(Count) = ""
            Add()
        Next
    End Sub

    Public Function IsEmpty() As Boolean
        If Rear = -1 Then
            Return True
        Else
            Return False
        End If
    End Function

    Public Function Remove() As Char
        Dim Item As Char
        If IsEmpty() Then
            Return Nothing
        Else
            Item = Contents(0)
            For Count = 1 To Rear
                Contents(Count - 1) = Contents(Count)
            Next
            Contents(Rear) = ""
            Rear -= 1
            Return Item
        End If
    End Function

    Public Sub Add()
        Dim RandNo As Integer
        If Rear < MaxSize - 1 Then
            RandNo = Int(Rnd() * 26)
            Rear += 1
            Contents(Rear) = Chr(65 + RandNo)
        End If
    End Sub

    Public Sub Show()
        If Rear <> -1 Then
            Console.WriteLine()
            Console.Write("The contents of the queue are: ")
            For Each Item In Contents
                Console.Write(Item)
            Next
            Console.WriteLine()
        End If
    End Sub
End Class

Module WordsWithAQA
    Sub Main()
        Dim MaxHandSize As Integer
        Dim MaxTilesPlayed As Integer
        Dim NoOfEndOfTurnTiles As Integer
        Dim StartHandSize As Integer
        Dim Choice As String
        Dim AllowedWords As New List(Of String)
        Dim TileDictionary As New Dictionary(Of Char, Integer)()
        Console.WriteLine("++++++++++++++++++++++++++++++++++++++")
        Console.WriteLine("+ Welcome to the WORDS WITH AQA game +")
        Console.WriteLine("++++++++++++++++++++++++++++++++++++++")
        Console.WriteLine()
        Console.WriteLine()
        LoadAllowedWords(AllowedWords)
        TileDictionary = CreateTileDictionary()
        MaxHandSize = 20
        MaxTilesPlayed = 50
        NoOfEndOfTurnTiles = 3
        StartHandSize = 15
        Choice = ""
        While Choice <> "9"
            DisplayMenu()
            Console.Write("Enter your choice: ")
            Choice = Console.ReadLine()
            If Choice = "1" Then
                PlayGame(AllowedWords, TileDictionary, True, StartHandSize, MaxHandSize, MaxTilesPlayed, NoOfEndOfTurnTiles)
            ElseIf Choice = "2" Then
                PlayGame(AllowedWords, TileDictionary, False, 15, MaxHandSize, MaxTilesPlayed, NoOfEndOfTurnTiles)
            End If
        End While
    End Sub

    Function CreateTileDictionary() As Dictionary(Of Char, Integer)
        Dim TileDictionary As New Dictionary(Of Char, Integer)()
        For Count = 0 To 25
            If Array.IndexOf({0, 4, 8, 13, 14, 17, 18, 19}, Count) > -1 Then
                TileDictionary.Add(Chr(65 + Count), 1)
            ElseIf Array.IndexOf({1, 2, 3, 6, 11, 12, 15, 20}, Count) > -1 Then
                TileDictionary.Add(Chr(65 + Count), 2)
            ElseIf Array.IndexOf({5, 7, 10, 21, 22, 24}, Count) > -1 Then
                TileDictionary.Add(Chr(65 + Count), 3)
            Else
                TileDictionary.Add(Chr(65 + Count), 5)
            End If
        Next
        Return TileDictionary
    End Function

    Sub DisplayTileValues(ByVal TileDictionary As Dictionary(Of Char, Integer), ByRef AllowedWords As List(Of String))
        Console.WriteLine()
        Console.WriteLine("TILE VALUES")
        Console.WriteLine()
        For Each Tile As KeyValuePair(Of Char, Integer) In TileDictionary
            Console.WriteLine("Points for " & Tile.Key & ": " & Tile.Value)
        Next
        Console.WriteLine()
    End Sub

    Function GetStartingHand(ByRef TileQueue As QueueOfTiles, ByVal StartHandSize As Integer) As String
        Dim Hand As String
        Hand = ""
        For Count = 0 To StartHandSize - 1
            Hand += TileQueue.Remove()
            TileQueue.Add()
        Next
        Return Hand
    End Function

    Sub LoadAllowedWords(ByRef AllowedWords As List(Of String))
        Try
            Dim FileReader As New System.IO.StreamReader("aqawords.txt")
            While FileReader.EndOfStream <> True
                AllowedWords.Add(FileReader.ReadLine().Trim().ToUpper())
            End While
            FileReader.Close()
        Catch
            AllowedWords.Clear()
        End Try
    End Sub

    Function CheckWordIsInTiles(ByVal Word As String, ByVal PlayerTiles As String) As Boolean
        Dim CopyOfTiles As String
        Dim InTiles As Boolean
        InTiles = True
        CopyOfTiles = PlayerTiles
        For Count = 0 To Len(Word) - 1
            If CopyOfTiles.Contains(Word(Count)) Then
                CopyOfTiles = Replace(CopyOfTiles, Word(Count), "", , 1)
            Else
                InTiles = False
            End If
        Next
        Return InTiles
    End Function

    Function CheckWordIsValid(ByVal Word As String, ByRef AllowedWords As List(Of String)) As Boolean
        Dim ValidWord As Boolean
        Dim Count As Integer
        Count = 0
        ValidWord = False
        While Count < AllowedWords.Count And Not ValidWord
            If AllowedWords(Count) = Word Then
                ValidWord = True
            End If
            Count += 1
        End While
        Return ValidWord
    End Function

    Sub AddEndOfTurnTiles(ByRef TileQueue As QueueOfTiles, ByRef PlayerTiles As String, ByVal NewTileChoice As String, ByVal Choice As String)
        Dim NoOfEndOfTurnTiles As Integer
        If NewTileChoice = "1" Then
            NoOfEndOfTurnTiles = Len(Choice)
        ElseIf NewTileChoice = "2" Then
            NoOfEndOfTurnTiles = 3
        Else
            NoOfEndOfTurnTiles = Len(Choice) + 3
        End If
        For Count = 0 To NoOfEndOfTurnTiles - 1
            PlayerTiles += TileQueue.Remove()
            TileQueue.Add()
        Next
    End Sub

    Sub FillHandWithTiles(ByRef TileQueue As QueueOfTiles, ByRef PlayerTiles As String, ByVal MaxHandSize As Integer)
        While Len(PlayerTiles) <= MaxHandSize
            PlayerTiles += TileQueue.Remove()
            TileQueue.Add()
        End While
    End Sub

    Function GetScoreForWord(ByVal Word As String, ByVal TileDictionary As Dictionary(Of Char, Integer)) As Integer
        Dim Score As Integer
        Score = 0
        For Count = 0 To Len(Word) - 1
            Score += TileDictionary(Word(Count))
        Next
        If Len(Word) > 7 Then
            Score += 20
        ElseIf Len(Word) > 5 Then
            Score += 5
        End If
        Return Score
    End Function

    Sub UpdateAfterAllowedWord(ByVal Word As String, ByRef PlayerTiles As String, ByRef PlayerScore As Integer, ByRef PlayerTilesPlayed As Integer, ByVal TileDictionary As Dictionary(Of Char, Integer), ByRef AllowedWords As List(Of String))
        PlayerTilesPlayed += Len(Word)
        For Each Letter In Word
            PlayerTiles = Replace(PlayerTiles, Letter, "", , 1)
        Next
        PlayerScore += GetScoreForWord(Word, TileDictionary)
    End Sub

    Sub UpdateScoreWithPenalty(ByRef PlayerScore As Integer, ByVal PlayerTiles As String, ByVal TileDictionary As Dictionary(Of Char, Integer))
        For Count = 0 To Len(PlayerTiles) - 1
            PlayerScore -= TileDictionary(PlayerTiles(Count))
        Next
    End Sub

    Function GetChoice()
        Dim Choice As String
        Console.WriteLine()
        Console.WriteLine("Either:")
        Console.WriteLine("     enter the word you would like to play OR")
        Console.WriteLine("     press 1 to display the letter values OR")
        Console.WriteLine("     press 4 to view the tile queue OR")
        Console.WriteLine("     press 7 to view your tiles again OR")
        Console.WriteLine("     press 0 to fill hand and stop the game.")
        Console.Write("> ")
        Choice = Console.ReadLine()
        Console.WriteLine()
        Choice = Choice.ToUpper()
        Return Choice
    End Function

    Function GetNewTileChoice()
        Dim NewTileChoice As String
        NewTileChoice = ""
        While Array.IndexOf({"1", "2", "3", "4"}, NewTileChoice) = -1
            Console.WriteLine("Do you want to:")
            Console.WriteLine("     replace the tiles you used (1) OR")
            Console.WriteLine("     get three extra tiles (2) OR")
            Console.WriteLine("     replace the tiles you used and get three extra tiles (3) OR")
            Console.WriteLine("     get no new tiles (4)?")
            Console.Write("> ")
            NewTileChoice = Console.ReadLine()
        End While
        Return NewTileChoice
    End Function

    Sub DisplayTilesInHand(ByVal PlayerTiles As String)
        Console.WriteLine()
        Console.WriteLine("Your current hand: " & PlayerTiles)
    End Sub

    Sub HaveTurn(ByVal PlayerName As String, ByRef PlayerTiles As String, ByRef PlayerTilesPlayed As String, ByRef PlayerScore As Integer, ByVal TileDictionary As Dictionary(Of Char, Integer), ByRef TileQueue As QueueOfTiles, ByRef AllowedWords As List(Of String), ByVal MaxHandSize As Integer, ByVal NoOfEndOfTurnTiles As Integer)
        Dim NewTileChoice As String
        Dim Choice As String
        Dim ValidChoice As Boolean
        Dim ValidWord As Boolean
        Console.WriteLine()
        Console.WriteLine(PlayerName & " it is your turn.")
        DisplayTilesInHand(PlayerTiles)
        NewTileChoice = "2"
        ValidChoice = False
        While Not ValidChoice
            Choice = GetChoice()
            If Choice = "1" Then
                DisplayTileValues(TileDictionary, AllowedWords)
            ElseIf Choice = "4" Then
                TileQueue.Show()
            ElseIf Choice = "7" Then
                DisplayTilesInHand(PlayerTiles)
            ElseIf Choice = "0" Then
                ValidChoice = True
                FillHandWithTiles(TileQueue, PlayerTiles, MaxHandSize)
            Else
                ValidChoice = True
                If Len(Choice) = 0 Then
                    ValidWord = False
                Else
                    ValidWord = CheckWordIsInTiles(Choice, PlayerTiles)
                End If
                If ValidWord Then
                    ValidWord = CheckWordIsValid(Choice, AllowedWords)
                    If ValidWord Then
                        Console.WriteLine()
                        Console.WriteLine("Valid word")
                        Console.WriteLine()
                        UpdateAfterAllowedWord(Choice, PlayerTiles, PlayerScore, PlayerTilesPlayed, TileDictionary, AllowedWords)
                        NewTileChoice = GetNewTileChoice()
                    End If
                End If
                If Not ValidWord Then
                    Console.WriteLine()
                    Console.WriteLine("Not a valid attempt, you lose your turn.")
                    Console.WriteLine()
                End If
                If NewTileChoice <> "4" Then
                    AddEndOfTurnTiles(TileQueue, PlayerTiles, NewTileChoice, Choice)
                End If
                Console.WriteLine()
                Console.WriteLine("Your word was: " & Choice)
                Console.WriteLine("Your new score is: " & PlayerScore)
                Console.WriteLine("You have played " & PlayerTilesPlayed & " tiles so far in this game.")
            End If
        End While
    End Sub

    Sub DisplayWinner(ByVal PlayerOneScore As Integer, ByVal PlayerTwoScore As Integer)
        Console.WriteLine()
        Console.WriteLine("**** GAME OVER! ****")
        Console.WriteLine()
        Console.WriteLine("Player One your score is " & PlayerOneScore)
        Console.WriteLine("Player Two your score is " & PlayerTwoScore)
        If PlayerOneScore > PlayerTwoScore Then
            Console.WriteLine("Player One wins!")
        ElseIf PlayerTwoScore > PlayerOneScore Then
            Console.WriteLine("Player Two wins!")
        Else
            Console.WriteLine("It is a draw!")
        End If
        Console.WriteLine()
    End Sub

    Sub PlayGame(ByRef AllowedWords As List(Of String), ByVal TileDictionary As Dictionary(Of Char, Integer), ByVal RandomStart As Boolean, ByVal StartHandSize As Integer, ByVal MaxHandSize As Integer, ByVal MaxTilesPlayed As Integer, ByVal NoOfEndOfTurnTiles As Integer)
        Dim PlayerOneScore As Integer
        Dim PlayerTwoScore As Integer
        Dim PlayerOneTilesPlayed As Integer
        Dim PlayerTwoTilesPlayed As Integer
        Dim PlayerOneTiles As String
        Dim PlayerTwoTiles As String
        Dim TileQueue As New QueueOfTiles(20)
        PlayerOneScore = 50
        PlayerTwoScore = 50
        PlayerOneTilesPlayed = 0
        PlayerTwoTilesPlayed = 0
        If RandomStart Then
            PlayerOneTiles = GetStartingHand(TileQueue, StartHandSize)
            PlayerTwoTiles = GetStartingHand(TileQueue, StartHandSize)
        Else
            PlayerOneTiles = "BTAHANDENONSARJ"
            PlayerTwoTiles = "CELZXIOTNESMUAA"
        End If
        While PlayerOneTilesPlayed <= MaxTilesPlayed And PlayerTwoTilesPlayed <= MaxTilesPlayed And Len(PlayerOneTiles) < MaxHandSize And Len(PlayerTwoTiles) < MaxHandSize
            HaveTurn("Player One", PlayerOneTiles, PlayerOneTilesPlayed, PlayerOneScore, TileDictionary, TileQueue, AllowedWords, MaxHandSize, NoOfEndOfTurnTiles)
            Console.WriteLine()
            Console.Write("Press Enter to continue")
            Console.ReadLine()
            Console.WriteLine()
            HaveTurn("Player Two", PlayerTwoTiles, PlayerTwoTilesPlayed, PlayerTwoScore, TileDictionary, TileQueue, AllowedWords, MaxHandSize, NoOfEndOfTurnTiles)
        End While
        UpdateScoreWithPenalty(PlayerOneScore, PlayerOneTiles, TileDictionary)
        UpdateScoreWithPenalty(PlayerTwoScore, PlayerTwoTiles, TileDictionary)
        DisplayWinner(PlayerOneScore, PlayerTwoScore)
    End Sub

    Sub DisplayMenu()
        Console.WriteLine()
        Console.WriteLine("=========")
        Console.WriteLine("MAIN MENU")
        Console.WriteLine("=========")
        Console.WriteLine()
        Console.WriteLine("1. Play game with random start hand")
        Console.WriteLine("2. Play game with training start hand")
        Console.WriteLine("9. Quit")
        Console.WriteLine()
    End Sub
End Module