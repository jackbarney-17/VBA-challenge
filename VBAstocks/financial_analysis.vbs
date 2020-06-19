Sub stockAnalysisLoop()
    
    ' setting dimension current as a worksheet object variable
    Dim Current As Worksheet

         ' for each statement to loop thru each worksheet in active workbook
        For Each Current In Worksheets

            ' naming dimensions (variables)
            Dim total As Double
            Dim start As Long
            Dim numRows As Long
            Dim days As Integer
            Dim change As Double
            Dim pctChange As Double
            Dim dailyChange As Double
            Dim avgChange As Double
            Dim i As Long
            Dim j As Integer
        
            ' naming titles
            Range("I1").Value = "Ticker"
            Range("J1").Value = "Yearly Change"
            Range("K1").Value = "Percent Change"
            Range("L1").Value = "Total Stock Volume"
        
            ' setting beginning values for counting
            start = 2
            change = 0
            total = 0
            j = 0
        
            ' counting rows up to last row containing data
            numRows = Cells(Rows.Count, "A").End(xlUp).Row
        
            For i = 2 To numRows
        
                ' print results when/if ticker changes value
                If Cells(i + 1, 1).Value <> Cells(i, 1).Value Then
        
                ' if/else statement to handle rows with 0 value in volume
                    If total = 0 Then
                        ' print results
                        Range("I" & 2 + j).Value = Cells(i, 1).Value
                        Range("J" & 2 + j).Value = 0
                        Range("K" & 2 + j).Value = "%" & 0
                        Range("L" & 2 + j).Value = 0
                    Else
                        ' Find first instance of non zero start value
                        If Cells(start, 3) = 0 Then
                            For get_value = start To i
                                If Cells(find_value, 3).Value <> 0 Then
                                    start = get_value
                                    Exit For
                                End If
                             Next get_value
                        End If
        
                        ' Get change/pct change
                        change = (Cells(i, 6) - Cells(start, 3))
                        pctChange = Round((change / Cells(start, 3) * 100), 2)
        
                        ' starting next ticker
                        start = i + 1
        
                        ' print results
                        Range("I" & 2 + j).Value = Cells(i, 1).Value
                        Range("J" & 2 + j).Value = Round(change, 2)
                        Range("K" & 2 + j).Value = "%" & pctChange
                        Range("L" & 2 + j).Value = total
        
                        ' making positive change values green and negative change values red
                        Select Case change
                            Case Is > 0
                                Range("J" & 2 + j).Interior.ColorIndex = 4
                            Case Is < 0
                                Range("J" & 2 + j).Interior.ColorIndex = 3
                            Case Else
                                Range("J" & 2 + j).Interior.ColorIndex = 0
                        End Select
                    End If
        
                    ' resetting our variables/dimensions for the next ticker
                    days = 0
                    change = 0
                    total = 0
                    j = j + 1
        
                ' add results if ticker is same as before
                Else
                    total = total + Cells(i, 7).Value
                End If
            Next i
            ' display current worksheet name in msg box for reference
            MsgBox Current.Name
        Next

      End Sub
