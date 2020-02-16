Sub Stocks()

    'Insert column titles in each worksheet
    
    Cells(1, 9).Value = "Ticker"
    Cells(1, 10).Value = "Yearly Change"
    Cells(1, 11).Value = "Percent Change"
    Cells(1, 12).Value = "Total Stock Volume"
    
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
    LastRow = Cells(Rows.Count, 1).End(xlUp).Row
    
    'Keep track of the location for each ticker in the summary table
    Dim SummaryTableRow As Integer
    SummaryTableRow = 2
  
    'Set initial value for Open Price
    OpenPrice = Cells(2, 3).Value

    'Loop through all ticker symbols
        For i = 2 To LastRow

            'Check if we are still within the same ticker symbol
                If Cells(i + 1, 1).Value <> Cells(i, 1).Value Then

                    'Set the ticker symbol
                    Ticker = Cells(i, 1).Value
                    Cells(SummaryTableRow, 9).Value = Ticker
                
                    ' Set closing price
                    ClosingPrice = Cells(i, 6).Value
               
                    'Add the yearly change
                    YearlyChange = ClosingPrice - OpenPrice
                    Cells(SummaryTableRow, 10).Value = YearlyChange
                
                        'Add the percent change
                        If (OpenPrice = 0 And ClosingPrice = 0) Then
                            PercentChange = 0
                
                        ElseIf (OpenPrice = 0 And ClosingPrice <> 0) Then
                                PercentChange = 1
                
                        Else
                            PercentChange = YearlyChange / OpenPrice
                            Cells(SummaryTableRow, 11).Value = PercentChange
                            
                            'Change the number format to %
                            Cells(SummaryTableRow, 11).NumberFormat = "0.00%"
               
                        End If
                
                    'Add the total volume
                    TotalVolume = TotalVolume + Cells(i, 7).Value
                    Cells(SummaryTableRow, 12).Value = TotalVolume
                
                    'Add one to the summary table row
                    SummaryTableRow = SummaryTableRow + 1
                
                    'Reset the open price
                    OpenPrice = Cells(i + 1, 3)
                
                    'Reset the total volume
                    TotalVolume = 0
            
                    'If the cells have the same ticker symbol
                    Else
                        TotalVolume = TotalVolume + Cells(i, 7).Value
            
                End If
        
            Next i
        
        'Determine the last row of the yearly change column
        YearlyChangeLastRow = Cells(Rows.Count, 9).End(xlUp).Row
        
        'Set the two cell colors
        'Positive change = green and Negative vhange = red
        For j = 2 To YearlyChangeLastRow
            If (Cells(j, 10).Value > 0 Or Cells(j, 11).Value = 0) Then
                Cells(j, 10).Interior.ColorIndex = 4
            
            ElseIf Cells(j, 10).Value < 0 Then
                Cells(j, 10).Interior.ColorIndex = 3
            
            End If
        
        Next j
End Sub
