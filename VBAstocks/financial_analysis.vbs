Sub analysis():

    ' Naming dimensions (variables)
    Dim total As Double
    Dim start As Long
    Dim numRows as Long
    Dim days as Integer
    Dim change as Double
    Dim pctChange as Double
    Dim dailyChange as Double
    Dim avgChange as Double
    Dim i as Long
    Dim j as Integer

    ' Naming titles
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

        ' if/else statement to handle rows with 0 value imn volume
            If total = 0 Then
                ' print results
                Range("I" & 2 + j).Value = Cells(i, 1).Value
                Range("J" & 2 + j).Value = 0
                Range("K" & 2 + j).Value = "%" & 0
                Range("L" & 2 + j).Value = 0
            Else
                ' Find first instance of non zero start value
                If Cells(start, 3) = 0 Then
                    For find_value = start To i
                        If Cells(find_value, 3).Value <> 0 Then
                            start = find_value
                            Exit For
                        End If
                     Next find_value
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
            total = 0
            change = 0
            j = j + 1
            days = 0

        ' add results if ticker is same as before
        Else
            total = total + Cells(i, 7).Value
        End If
    Next i
End Sub