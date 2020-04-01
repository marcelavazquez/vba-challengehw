Attribute VB_Name = "Module1"
Sub stocks()

Dim ws As Worksheet

'Walk through each worksheet
    For Each ws In ActiveWorkbook.Worksheets
    ws.Activate

        'Declare variables
        Dim ticker As String
        Dim open_price As Double
        Dim close_price As Double
        Dim yearlyChange As Double
        Dim percentChange As Double

        Dim volume As Double
        volume = 0

        rowdelimitier = Cells(Rows.Count, 1).End(xlUp).Row
 
        open_price = Cells(2, 3).Value

        'Store the infromation needed in the summary table
        Dim Summary_Table_Row As Integer
        Summary_Table_Row = 2

        'Give headers to summary table
        Range("I1").Value = "Ticker"
        Range("J1").Value = "Yearly Change"
        Range("K1").Value = "Percent Change"
        Range("L1").Value = "Total Stock Volume"

        'Iterate through each row and column
            For i = 2 To rowdelimitier
        
                'Check if we are still within the same ticker, if it is not...
                If Cells(i + 1, 1).Value <> Cells(i, 1).Value Then
    
                'Set the Ticker, open_price, close_price, yearlyChange, percentageChange and volume values
                ticker = Cells(i, 1).Value
                
                close_price = Cells(i, 6).Value
                
                yearlyChange = close_price - open_price
        
                If open_price = 0 Then
                percentChange = 0
                Else
                percentChange = yearlyChange / open_price
                End If
    
                Range("K" & Summary_Table_Row).NumberFormat = "0.00%"
                
                'Add to the total stock volume
                volume = volume + Cells(i, 7).Value
    
                'Print the ticker , yearly change, percentage change and total stock volume in the summary table
                Range("I" & Summary_Table_Row).Value = ticker
                Range("J" & Summary_Table_Row).Value = yearlyChange
                Range("K" & Summary_Table_Row).Value = percentChange
                Range("L" & Summary_Table_Row).Value = volume
    
                  'Highlight positive change in green and negative change in red.
                If yearlyChange > 0 Then
                    Range("J" & Summary_Table_Row).Interior.ColorIndex = 4
                ElseIf yearlyChange < 0 Then
                    Range("J" & Summary_Table_Row).Interior.ColorIndex = 3
                End If
    
                'Add one to the summary table row
                Summary_Table_Row = Summary_Table_Row + 1
    
                'Reset volume
                volume = 0
    
                'Reset open price
                open_price = Cells(i + 1, 3)
    
                'if the cell immediately following a row in the same ticker
                Else
        
                'Add to the volume
                volume = volume + Cells(i, 7).Value
    
                End If

            Next i
            
            'Challenge headers
            Range("O2").Value = "Greatest % Increase"
            Range("O3").Value = "Greatest % Decrease"
            Range("O4").Value = "Greatest Total Volume"
            Range("P1").Value = "Ticker"
            Range("Q1").Value = "Value"
            
            'Bonus determine variables
            Dim gr_percent_increase As Double
            Dim ticker_increase As String
            Dim gr_percent_decrease As Double
            Dim ticker_decrease As String
            Dim gr_total_volume As Double
            Dim ticker_volume As String
            
            'Determine new row delimitier for  Yearly Change
            rowdelimitierbonus = Cells(Rows.Count, 9).End(xlUp).Row
            
            'Set variables
            gr_percent_increase = 0
            gr_percent_decrease = 0
            gr_total_volume = 0
            
            'find the greatest value and their corresponding ticker names
            For j = 2 To rowdelimitierbonus
            
                If Cells(j, 11).Value > gr_percent_increase Then
                    gr_percent_increase = Cells(j, 11).Value
                    ticker_increase = Cells(j, 9).Value
                End If
                
                If Cells(j, 11).Value < gr_percent_decrease Then
                    gr_percent_decrease = Cells(j, 11).Value
                    ticker_decrease = Cells(j, 9).Value
                End If
                
                If Cells(j, 12).Value > gr_total_volume Then
                    gr_total_volume = Cells(j, 12).Value
                    ticker_volume = Cells(j, 9).Value
                End If
                
             Next j
                
            
            'Print data in second table
            Range("P2").Value = ticker_increase
            Range("Q2").Value = gr_percent_increase
            Range("Q2").NumberFormat = "0.00%"
            
            Range("P3").Value = ticker_decrease
            Range("Q3").Value = gr_percent_decrease
            Range("Q3").NumberFormat = "0.00%"
            
            Range("P4") = ticker_volume
            Range("Q4").Value = gr_total_volume
            
    Next ws

End Sub
