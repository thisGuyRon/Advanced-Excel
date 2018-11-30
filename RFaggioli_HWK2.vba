Sub Faggioli_HW2()
        'Variable declaration
    Dim i As Long
    Dim j As Long
    Dim Ticker As String
    Dim initValue As Date
    Dim finValue As Date
    Dim volume As LongLong
    Dim rowCount As Long
    Dim decTicker As String
    Dim incTicker As String
    Dim incValue As Double
    Dim decValue As Double
    Dim volTicker As String
    Dim volValue As LongLong
    
    
    'print column Names
    'moderate columns
    Cells(1, 9) = "Ticker"
    Cells(1, 10) = "Yearly Price Change"
    Cells(1, 11) = "% Yearly Price Change"
    Cells(1, 12) = "Total Volume"
    'hard columns
    Cells(1, 15) = "Ticker"
    Cells(1, 16) = "Value"
    Cells(2, 14) = "Greatest % Increase"
    Cells(3, 14) = "Greatest % Decrease"
    Cells(4, 14) = "Greatest Total Volume"
    
    'to capture the row count
    rowCount = Range("A1", Range("A1").End(xlDown)).Rows.Count
    
    'Initial set of variables
    Ticker = Cells(2, 1)
    initValue = Cells(2, 3)
    volume = 0
    j = 2
    
    For i = 2 To rowCount + 1
        'Check for ticker value change if so prints values to excel
        If Cells(i, 1) <> Ticker Then
                Cells(j, 9) = Ticker
                Cells(j, 10) = finValue - initValue
                
                'catches divisible by 0 error
                If finValue = 0 Then
                    Cells(j, 11) = 0
                Else
                    Cells(j, 11) = (Cells(j, 10) / finValue)
                End If
                
                'conditional formatting
                If Cells(j, 10) < 0 Then
                    Cells(j, 10).Interior.ColorIndex = 3
                 ElseIf Cells(j, 10) > 0 Then
                    Cells(j, 10).Interior.ColorIndex = 4
                End If
                If Cells(j, 11) < 0 Then
                    Cells(j, 11).Interior.ColorIndex = 3
                 ElseIf Cells(j, 11) > 0 Then
                    Cells(j, 11).Interior.ColorIndex = 4
                End If
                
                Cells(j, 12) = volume
                
                'increments j for next line of ticker results
                j = j + 1
                'sets new base values for next ticker results
                Ticker = Cells(i, 1)
                initValue = Cells(i, 3)
                finValue = Cells(i, 6)
                volume = Cells(i, 7)
                
            'updates final ticker price and running total of volume if ticker matches
        ElseIf Cells(i, 1) = Ticker Then
                finValue = Cells(i, 6)
                volume = volume + Cells(i, 7)
        End If
Next i

'Hard Code
'to capture the new row count
    rowCount = Range("I2", Range("I1").End(xlDown)).Rows.Count

    
    For i = 2 To rowCount
        'set Variable for greatest Increase
        If Cells(i, 11) > incValue Then
            incValue = Cells(i, 11)
            incTicker = Cells(i, 9)
        End If
        
        'Set Variable for greatest decrease
        If Cells(i, 11) < decValue Then
            decValue = Cells(i, 11)
            decTicker = Cells(i, 9)
        End If
        
        'Set for greatest Volume
        If Cells(i, 12) > volValue Then
            volValue = Cells(i, 12)
            volTicker = Cells(i, 9)
        End If
    Next i
    'writing hard fields
    Cells(2, 15) = incTicker
    Cells(3, 15) = decTicker
    Cells(4, 15) = volTicker
    Cells(2, 16) = incValue
    Cells(3, 16) = decValue
    Cells(4, 16) = volValue

End Sub
