Dim board As Range
Private Function smartPlay() As Range
    Set smartPlay = winPlay
    If smartPlay.cells.Count < 9 Then Exit Function
    Set smartPlay = blockWin
    If smartPlay.cells.Count < 9 Then Exit Function
    Set smartPlay = bifurcate
    If smartPlay.cells.Count < 9 Then Exit Function
    Set smartPlay = blockBif
End Function
Private Function winPlay() As Range
    For i = 1 To 3:
        If WorksheetFunction.CountIf(board.Columns(i), "o") = 2 And WorksheetFunction.CountIf(board.Columns(i), "") = 1 Then
            Set winPlay = board.Columns(i)
            Exit Function
        End If
        If WorksheetFunction.CountIf(board.Rows(i), "o") = 2 And WorksheetFunction.CountIf(board.Rows(i), "") = 1 Then
            Set winPlay = board.Rows(i)
            Exit Function
        End If
    Next i
    If -(board.Range("A1").Value = "o") - (board.Range("B2").Value = "o") - (board.Range("C3") = "o") = 2 And -(board.Range("A1").Value = "") - (board.Range("B2").Value = "") - (board.Range("C3") = "") = 1 Then
        Set winPlay = board.Range("A1,B2,C3")
        Exit Function
    End If
    If -(board.Range("C1").Value = "o") - (board.Range("B2").Value = "o") - (board.Range("A3") = "o") = 2 And -(board.Range("C1").Value = "") - (board.Range("B2").Value = "") - (board.Range("A3") = "") = 1 Then
        Set winPlay = board.Range("C1,B2,A3")
        Exit Function
    End If
    Set winPlay = board
End Function
Private Function blockWin() As Range
    For i = 1 To 3:
        If WorksheetFunction.CountIf(board.Columns(i), "x") = 2 And WorksheetFunction.CountIf(board.Columns(i), "") = 1 Then
            Set blockWin = board.Columns(i)
            Exit Function
        End If
        If WorksheetFunction.CountIf(board.Rows(i), "x") = 2 And WorksheetFunction.CountIf(board.Rows(i), "") = 1 Then
            Set blockWin = board.Rows(i)
            Exit Function
        End If
    Next i
    If -(board.Range("A1").Value = "x") - (board.Range("B2").Value = "x") - (board.Range("C3") = "x") = 2 And -(board.Range("A1").Value = "") - (board.Range("B2").Value = "") - (board.Range("C3") = "") = 1 Then
        Set blockWin = board.Range("A1,B2,C3")
        Exit Function
    End If
    If -(board.Range("C1").Value = "x") - (board.Range("B2").Value = "x") - (board.Range("A3") = "x") = 2 And -(board.Range("C1").Value = "") - (board.Range("B2").Value = "") - (board.Range("A3") = "") = 1 Then
        Set blockWin = board.Range("C1,B2,A3")
        Exit Function
    End If
    Set blockWin = board
End Function
Private Function bifurcate() As Range
    Set bifurcate = board.cells(0, 0)
    For i = 1 To 3:
        If WorksheetFunction.CountIf(board.Columns(i), "o") = 1 And WorksheetFunction.CountIf(board.Columns(i), "") = 2 Then
            For Each c In board.Columns(i).cells
                If Not Application.Intersect(bifurcate, c) Is Nothing Then
                    Set bifurcate = c
                    Exit Function
                End If
                If c.Value = "" Then Set bifurcate = Union(bifurcate, c)
            Next c
        End If
        If WorksheetFunction.CountIf(board.Rows(i), "o") = 1 And WorksheetFunction.CountIf(board.Rows(i), "") = 2 Then
            For Each c In board.Rows(i).cells
                If Not Application.Intersect(bifurcate, c) Is Nothing Then
                    Set bifurcate = c
                    Exit Function
                End If
                If c.Value = "" Then Set bifurcate = Union(bifurcate.cells, c)
            Next c
        End If
    Next i
    If -(board.Range("A1").Value = "o") - (board.Range("B2").Value = "o") - (board.Range("C3") = "o") = 1 And -(board.Range("A1").Value = "") - (board.Range("B2").Value = "") - (board.Range("C3") = "") = 2 Then
        For Each c In board.Range("A1,B2,C3")
                If Not Application.Intersect(bifurcate, c) Is Nothing Then
                    Set bifurcate = c
                    Exit Function
                End If
                If c.Value = "" Then Set bifurcate = Union(bifurcate, c)
        Next c
    End If
    If -(board.Range("C1").Value = "o") - (board.Range("B2").Value = "o") - (board.Range("A3") = "o") = 1 And -(board.Range("C1").Value = "") - (board.Range("B2").Value = "") - (board.Range("A3") = "") = 2 Then
        For Each c In board.Range("C1,B2,A3")
            If Not Application.Intersect(bifurcate, c) Is Nothing Then
                Set bifurcate = c
                Exit Function
            End If
            If c.Value = "" Then Set bifurcate = Union(bifurcate, c)
        Next c
    End If
    Set bifurcate = board
End Function
Private Function blockBif()
    Set blockBif = board.cells(0, 0)
    For i = 1 To 3:
        If WorksheetFunction.CountIf(board.Columns(i), "x") = 1 And WorksheetFunction.CountIf(board.Columns(i), "") = 2 Then
            For Each c In board.Columns(i).cells
                If Not Application.Intersect(blockBif, c) Is Nothing Then
                    Set blockBif = c
                    Exit Function
                End If
                If c.Value = "" Then Set blockBif = Union(blockBif, c)
            Next c
        End If
        If WorksheetFunction.CountIf(board.Rows(i), "x") = 1 And WorksheetFunction.CountIf(board.Rows(i), "") = 2 Then
            For Each c In board.Rows(i).cells
                If Not Application.Intersect(blockBif, c) Is Nothing Then
                    Set blockBif = c
                    Exit Function
                End If
                If c.Value = "" Then Set blockBif = Union(blockBif.cells, c)
            Next c
        End If
    Next i
    If -(board.Range("A1").Value = "x") - (board.Range("B2").Value = "x") - (board.Range("C3") = "x") = 1 And -(board.Range("A1").Value = "") - (board.Range("B2").Value = "") - (board.Range("C3") = "") = 2 Then
        For Each c In board.Range("A1,B2,C3")
                If Not Application.Intersect(blockBif, c) Is Nothing Then
                    Set blockBif = c
                    Exit Function
                End If
                If c.Value = "" Then Set blockBif = Union(blockBif, c)
        Next c
    End If
    If -(board.Range("C1").Value = "x") - (board.Range("B2").Value = "x") - (board.Range("A3") = "x") = 1 And -(board.Range("C1").Value = "") - (board.Range("B2").Value = "") - (board.Range("A3") = "") = 2 Then
        For Each c In board.Range("C1,B2,A3")
            If Not Application.Intersect(blockBif, c) Is Nothing Then
                Set blockBif = c
                Exit Function
            End If
            If c.Value = "" Then Set blockBif = Union(blockBif, c)
        Next c
    End If
    Set blockBif = board
End Function
Private Function botPlay() As Range
    If WorksheetFunction.CountIf(board, "") <= 6 Then
        Set botPlay = smartPlay
        If botPlay.cells.Count < 9 Then
            For Each c In botPlay.cells
                If c.Value = "" Then
                    Set botPlay = c
                    Exit Function
                End If
            Next c
        End If
    End If
    Dim slot, freeSlots As Integer
    freeSlots = WorksheetFunction.CountIf(board, "")
    slot = freeSlots * Rnd
    For Each c In board.cells:
        If c.Value = "" Then
            Set botPlay = c
            slot = slot - 1
            If slot < 0 Then Exit For
        End If
    Next c
End Function
Private Function checkWin() As Boolean
    checkWin = True
    For i = 1 To 3:
        If WorksheetFunction.CountIf(board.Columns(i), "x") = 3 Then
            board.Range("A6").Value = board.Range("A6").Value + 1
            Call colorRange(board.Columns(i), RGB(0, 255, 0))
            Exit Function
        End If
        If WorksheetFunction.CountIf(board.Rows(i), "x") = 3 Then
            board.Range("A6").Value = board.Range("A6").Value + 1
            Call colorRange(board.Rows(i), RGB(0, 255, 0))
            Exit Function
        End If
    Next i
    If board.Range("A1").Value = board.Range("B2") And board.Range("B2") = board.Range("C3") And board.Range("A1") = "x" Then
        board.Range("A6").Value = board.Range("A6").Value + 1
        Call colorRange(board.Range("A1,B2,C3"), RGB(0, 255, 0))
        Exit Function
    End If
    If board.Range("C1").Value = board.Range("B2") And board.Range("B2") = board.Range("A3") And board.Range("C1") = "x" Then
        board.Range("A6").Value = board.Range("A6").Value + 1
        Call colorRange(board.Range("C1,B2,A3"), RGB(0, 255, 0))
        Exit Function
    End If
    For i = 1 To 3:
        If WorksheetFunction.CountIf(board.Columns(i), "o") = 3 Then
            board.Range("C6").Value = board.Range("C6").Value + 1
            Call colorRange(board.Columns(i), RGB(255, 0, 0))
            Exit Function
        End If
        If WorksheetFunction.CountIf(board.Rows(i), "o") = 3 Then
            board.Range("C6").Value = board.Range("C6").Value + 1
            Call colorRange(board.Rows(i), RGB(255, 0, 0))
            Exit Function
        End If
    Next i
    If board.Range("A1").Value = board.Range("B2") And board.Range("B2") = board.Range("C3") And board.Range("A1") = "o" Then
        board.Range("C6").Value = board.Range("C6").Value + 1
        Call colorRange(board.Range("A1,B2,C3"), RGB(255, 0, 0))
        Exit Function
    End If
    If board.Range("C1").Value = board.Range("B2") And board.Range("B2") = board.Range("A3") And board.Range("C1") = "o" Then
        board.Range("C6").Value = board.Range("C6").Value + 1
        Call colorRange(board.Range("C1,B2,A3"), RGB(255, 0, 0))
        Exit Function
    End If
    If WorksheetFunction.CountIf(board, "") = 0 Then
        board.Range("B6").Value = board.Range("B6").Value + 1
        Call colorRange(board, RGB(255, 255, 0))
        Exit Function
    End If
    checkWin = False
End Function
Sub colorRange(cells As Range, clr As Long)
    cells.Interior.Color = clr
End Sub
Sub newGame()
    board.Value = ""
    Application.Wait (Now + TimeValue("0:0:1"))
    Call colorRange(board, RGB(255, 255, 255))
    If Rnd >= 0.5 Then botPlay.Value = "o"
End Sub
Sub mainClick(tile As Range)
    Call defineBoard
    If tile.Value = "" Then
        tile.Value = "x"
        If checkWin Then
            Call newGame
            Exit Sub
        End If
        botPlay.Value = "o"
        If checkWin Then Call newGame
    End If
End Sub
Sub click1()
    Call mainClick(Plan1.Range("B2"))
End Sub
Sub click2()
    Call mainClick(Plan1.Range("B3"))
End Sub
Sub click3()
    Call mainClick(Plan1.Range("B4"))
End Sub
Sub click4()
    Call mainClick(Plan1.Range("C2"))
End Sub
Sub click5()
    Call mainClick(Plan1.Range("C3"))
End Sub
Sub click6()
    Call mainClick(Plan1.Range("C4"))
End Sub
Sub click7()
    Call mainClick(Plan1.Range("D2"))
End Sub
Sub click8()
    Call mainClick(Plan1.Range("D3"))
End Sub
Sub click9()
    Call mainClick(Plan1.Range("D4"))
End Sub
Sub defineBoard()
    Set board = Plan1.Range("B2:D4")
End Sub
Sub clearScore()
    Call defineBoard
    board.Range("A6:C6").Value = 0
    Call newGame
End Sub
