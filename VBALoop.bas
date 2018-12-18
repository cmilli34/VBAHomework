Attribute VB_Name = "Module1"
Sub looping()
    Dim i As Long
    Dim previousTicker As String
    Dim currentTicker As String
    Dim volume As Double
    Dim tickerDisplayRow As Integer
    
    previousTicker = Cells(2, 1).Value
    volume = 0
    tickerDisplayRow = 1
    
    Cells(1, 9).Value = "TIcker"
    Cells(1, 10).Value = "Total Volume"
   
    
    Last = Cells(Rows.Count, 1).End(xlUp).Row
    
    For i = 2 To Last
        If Cells(i, 1).Value <> "" Then
            currentTicker = Cells(i, 1).Value
            If previousTicker <> currentTicker Then
                tickerDisplayRow = tickerDisplayRow + 1
                Cells(tickerDisplayRow, 9).Value = previousTicker
                Cells(tickerDisplayRow, 10).Value = volume
                volume = 0
            End If
            volume = volume + Cells(i, 7).Value
            previousTicker = currentTicker
        Else
            Exit For
        End If
    Next i
    
    ' Print out the last line
    tickerDisplayRow = tickerDisplayRow + 1
    Cells(tickerDisplayRow, 9).Value = previousTicker
    Cells(tickerDisplayRow, 10).Value = volume
    
    MsgBox ("It's a Miracle")
End Sub
