Sub Stocks()

    For Each WS In Worksheets
    
    'Insert column titles in each worksheet
    
    WS.Cells(1, 9).Value = "Ticker"
    WS.Cells(1, 10).Value = "Yearly Change"
    WS.Cells(1, 11).Value = "Percent Change"
    WS.Cells(1, 12).Value = "Total Stock Volume"
    
    'Define variables to hold values
     
    Dim Ticker As String
    Dim OpenPrice As Double
    Dim ClosingPrice As Double
    Dim YearlyChange As Double
    Dim PercentChange As Double
    
    
    'Set an initial variable for holding the total stock volume
    Dim TotalVolume As Double
    TotalVolume = 0
    
    'Determine the last row
    LastRow = WS.Cells(Rows.Count, 1).End(xlUp).Row
    
    'Keep track of the location for each ticker in the summary table
    Dim SummaryTableRow As Integer
    SummaryTableRow = 2
  
    'Set initial value for Open Price
    OpenPrice = WS.Cells(2, 3).Value

    'Loop through all ticker symbols
        For i = 2 To LastRow

            'Check if we are still within the same ticker symbol
                If WS.Cells(i + 1, 1).Value <> WS.Cells(i, 1).Value Then

                    'Set the ticker symbol
                    Ticker = WS.Cells(i, 1).Value
                    WS.Cells(SummaryTableRow, 9).Value = Ticker
                
                    ' Set closing price
                    ClosingPrice = WS.Cells(i, 6).Value
               
                    'Add the yearly change
                    YearlyChange = ClosingPrice - OpenPrice
                    WS.Cells(SummaryTableRow, 10).Value = YearlyChange
                
                        'Add the percent change
                        If (OpenPrice = 0 And ClosingPrice = 0) Then
                            PercentChange = 0
                
                        ElseIf (OpenPrice = 0 And ClosingPrice <> 0) Then
                                PercentChange = 1
                
                        Else
                            PercentChange = YearlyChange / OpenPrice
                            WS.Cells(SummaryTableRow, 11).Value = PercentChange
                            
                            'Change the number format to %
                            WS.Cells(SummaryTableRow, 11).NumberFormat = "0.00%"
               
                        End If
                
                    'Add the total volume
                    TotalVolume = TotalVolume + WS.Cells(i, 7).Value
                    WS.Cells(SummaryTableRow, 12).Value = TotalVolume
                
                    'Add one to the summary table row
                    SummaryTableRow = SummaryTableRow + 1
                
                    'Reset the open price
                    OpenPrice = WS.Cells(i + 1, 3)
                
                    'Reset the total volume
                    TotalVolume = 0
            
                    'If the cells have the same ticker symbol
                    Else
                        TotalVolume = TotalVolume + WS.Cells(i, 7).Value
            
                End If
        
            Next i
        
        'Determine the last row of the yearly change column
        YearlyChangeLastRow = WS.Cells(Rows.Count, 9).End(xlUp).Row
        
        'Set the two cell colors
        'Positive change = green and Negative vhange = red
        For j = 2 To YearlyChangeLastRow
            If (WS.Cells(j, 10).Value > 0 Or WS.Cells(j, 11).Value = 0) Then
                WS.Cells(j, 10).Interior.ColorIndex = 4
            
            ElseIf WS.Cells(j, 10).Value < 0 Then
                WS.Cells(j, 10).Interior.ColorIndex = 3
            
            End If
        
        Next j

    'Set Greatest % Increase, Greatest % Decrease, and Greatest Total Volume
        
    WS.Cells(2, 15).Value = "Greatest % Increase"
    WS.Cells(3, 15).Value = "Greatest % Decrease"
    WS.Cells(4, 15).Value = "Greatest Total Volume"
    WS.Cells(1, 16).Value = "Ticker"
    WS.Cells(1, 17).Value = "Value"
        
    'Loop through each row to locate the greatest value and its ticker symbol
        For k = 2 To YearlyChangeLastRow
            
            If WS.Cells(k, 11).Value = Application.WorksheetFunction.Max(WS.Range("K2:K" & YearlyChangeLastRow)) Then
                WS.Cells(2, 16).Value = WS.Cells(k, 9).Value
                WS.Cells(2, 17).Value = WS.Cells(k, 11).Value
                
                'Change the number format to %
                WS.Cells(2, 17).NumberFormat = "0.00%"
            
            ElseIf WS.Cells(k, 11).Value = Application.WorksheetFunction.Min(WS.Range("K2:K" & YearlyChangeLastRow)) Then
                WS.Cells(3, 16).Value = WS.Cells(k, 9).Value
                WS.Cells(3, 17).Value = WS.Cells(k, 11).Value
                
                'Change the number format to %
                WS.Cells(3, 17).NumberFormat = "0.00%"
            
            ElseIf WS.Cells(k, 12).Value = Application.WorksheetFunction.Max(WS.Range("L2:L" & YearlyChangeLastRow)) Then
                WS.Cells(4, 16).Value = WS.Cells(k, 9).Value
                WS.Cells(4, 17).Value = WS.Cells(k, 12).Value
            
            End If
        
        Next k
        
    Next WS
        
End Sub


      