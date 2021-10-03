Attribute VB_Name = "Module1"
Sub move(col_change, row_change)
    
    If Worksheets("View").Range("D11").Value = 0 Then Exit Sub
    
    currentRowent_col = Range("E3").Value
    currentRowent_row = Range("F4").Value
    
    new_col = currentRowent_col + col_change
    new_row = currentRowent_row + row_change
    
    maze_value = Range("mazegrid").Cells(new_row, new_col).Value

    If maze_value = "" Or IsNumeric(maze_value) = True Then
        Range("E3").Value = new_col
        Range("F4").Value = new_row
        
        If maze_value <> "" And maze_value <> 1000 Then
            MsgBox "Secret Code: " & maze_value, vbInformation
            Range("mazegrid").Cells(new_row, new_col).Value = 1000
        End If
    Else
        Worksheets("View").Range("A8") = Worksheets("View").Range("A8") - 1
    End If
    
End Sub

Sub up()
    move 0, -1
End Sub

Sub down()
    move 0, 1
End Sub

Sub left()
    move -1, 0
End Sub

Sub right()
    move 1, 0
End Sub

Sub Build(sq)
    
    'Number of treasure points
    treasure = sq ^ 0.6
    treasuresum = 0
    
    'maxmoves assigned to maze
    Dim maxmoves As Integer
    maxmoves = 0
    
    'To count how many moves have been taken
    Dim move As Integer
    
    ' Define the number of attempts it will make to fill maxmoves before terminating
    term = (sq * sq) / 2
    
    'To save current location
    Dim location(0 To 100, 0 To 100)
    
    'To consider the four directions of possible travel
    Dim direction(0 To 3)
    
    'To record the route used to get to current location
    Dim route() As Variant
    
    'However as this is reset as the route changes, final route needs saving
    Dim final() As Variant
    
    'Redim to define to include 2nd dimension
    'route(0,move) = row
    'route(1,move) = col
    ReDim route(1, 0)
    ReDim final(1, 0)
    
    'Create a random starting point
    startR = Int(Rnd * (sq - 2)) + 1
    startC = Int(Rnd * (sq - 2)) + 1
    
    Range("F4").Value = startR
    Range("E3").Value = startC
    Range("A8").Value = 3
    
    Application.ScreenUpdating = False
    ThisWorkbook.Worksheets("Random Maze").Activate
    
    'Save value 1 to each cell in maze
    Range(Cells(1, 1), Cells(sq, sq)) = "w"
     
     
    'Put an S in the starting cell
    Cells(startR, startC) = "S"
     
    'Save starting point as first location and point on route
    location(startR, startC) = 1
    route(0, 0) = startR
    route(1, 0) = startC
    final(1, 0) = startR
    final(1, 0) = startC
    
    Do
        C = 0
        For i = 0 To 3
            'initally set all directions to be impossible
            direction(i) = 0
        Next i
        
        variable = UBound(route, 1)
        move = UBound(route, 2)
        
        currentRow = route(0, move)
        currentColumn = route(1, move)
        
        'Check up
        If currentRow - 2 >= 1 Then
            If location(currentRow - 1, currentColumn) <> 1 And _
                location(currentRow - 2, currentColumn) <> 1 And _
                location(currentRow - 1, currentColumn + 1) <> 1 And _
                location(currentRow - 1, currentColumn - 1) <> 1 Then
                direction(0) = 1
                C = 1
            End If
        End If
        
        'Check down
        If currentRow + 2 <= sq Then
            If location(currentRow + 1, currentColumn) <> 1 And _
               location(currentRow + 2, currentColumn) <> 1 And _
               location(currentRow + 1, currentColumn + 1) <> 1 And _
               location(currentRow + 1, currentColumn - 1) <> 1 Then
               direction(1) = 1
               C = 1
            End If
        End If
     
        'Check left
        If currentColumn - 2 >= 1 Then
            If location(currentRow, currentColumn - 1) <> 1 And _
               location(currentRow, currentColumn - 2) <> 1 And _
               location(currentRow + 1, currentColumn - 1) <> 1 And _
               location(currentRow - 1, currentColumn - 1) <> 1 Then
               direction(2) = 1
               C = 1
            End If
        End If
     
        'Check right
        If currentColumn + 2 <= sq Then
            If location(currentRow, currentColumn + 1) <> 1 And _
               location(currentRow, currentColumn + 2) <> 1 And _
               location(currentRow + 1, currentColumn + 1) <> 1 And _
               location(currentRow - 1, currentColumn + 1) <> 1 Then
               direction(3) = 1
               C = 1
            End If
        End If
     
        'If there are no possible moves
        If C = 0 Then
        
            If move > maxmoves Then
                maxmoves = move
                endR = route(0, UBound(route, 2))
                endC = route(1, UBound(route, 2))
                
                ReDim final(1, maxmoves)
                For i = 0 To maxmoves
                    final(0, i) = route(0, i)
                    final(1, i) = route(1, i)
                Next i
                
            End If
            
            If move = 0 Then
                Exit Do
            Else
                'Removes the last move to try a different route
                ReDim Preserve route(variable, move - 1)
            End If
            
            
            
         
        Else
            C = 0
            
            Do Until C <> 0
                'Randomly select direction direction
                rrand = Int(Rnd * 4)
                If direction(rrand) = 1 Then C = rrand + 1
            Loop
             
                
            'Add new move
            new_move = move + 1
            ReDim Preserve route(variable, new_move)
            
            
            Select Case C
                Case 1 'up
                    route(0, new_move) = currentRow - 1
                    route(1, new_move) = currentColumn
                    
                Case 2 'down
                    route(0, new_move) = currentRow + 1
                    route(1, new_move) = currentColumn
                    
                Case 3 'left
                    route(0, new_move) = currentRow
                    route(1, new_move) = currentColumn - 1
                    
                Case 4 'right
                    route(0, new_move) = currentRow
                    route(1, new_move) = currentColumn + 1
            End Select
             
            newR = route(0, new_move)
            newC = route(1, new_move)
             
            Range(Cells(1, 1), Cells(sq, sq)).Cells(newR, newC).Value = ""
             
            'Update currentRowent
            location(newR, newC) = 1
             
            'Apply changes
            DoEvents
        End If
        
    Loop
    
    'Add treasure
    For i = 1 To treasure
        treasuremove = i * (maxmoves \ (treasure))
        treasurevalue = 100 + Int(900 * Rnd)
        Cells(final(0, treasuremove), final(1, treasuremove)) = treasurevalue
        treasuresum = treasuresum + treasurevalue
    Next i
     
    'Add E to end of maze
    Cells(endR, endC) = "E"
    
    'Set named range
    ThisWorkbook.Names.Add Name:="mazegrid", RefersTo:=Range(Cells(1, 1), Cells(sq, sq))
    
    ThisWorkbook.Worksheets("View").Activate
    
    Range("G8").Value = treasuresum
    
    Application.ScreenUpdating = True


End Sub

Sub countdown()
    interval = Now + TimeValue("00:00:01")
    
    'If time is up
    If Worksheets("View").Range("D11").Value = 0 Then Exit Sub
    
    'If no lives
    If Worksheets("View").Range("A8").Value = 0 Then
        Worksheets("View").Range("D11").Value = 0
        Exit Sub
    End If

    Worksheets("View").Range("D11") = Worksheets("View").Range("D11") - TimeValue("00:00:01")
    
    Application.OnTime interval, "countdown"
    
End Sub


Sub resetmaze()
    'Feed in the size of the grid and create maze
    sq = Range("A9").Value
    Build sq
    Range("D11") = Format((sq \ 5) * TimeValue("00:00:35"), "hh:mm:ss")
    Call countdown
End Sub

    



