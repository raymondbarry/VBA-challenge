' ## Instructions
Sub Stock_VBA()
' * Create a script that will loop through all the stocks for one year for each run and take the following information.

    For Each ws In Worksheets
    '   * The ticker symbol.
        Dim Ticker_symbol As String
    '   * Yearly change from opening price at the beginning of a given year to the closing price at the end of that year.
        Dim Year_Change As Double
    ' '   * The percent change from opening price at the beginning of a given year to the closing price at the end of that year.
        Dim YearPChange As Integer
        Dim Open_Price As Double
        Dim Close_Price As Double
    '   * The total stock volume of the stock.
        Dim volume As Integer
        Dim volume_total As Double
        Dim summary_table_row As Integer
        Dim i As Long
        
        Dim LastRow As Long
        LastRow = Cells(Rows.Count, 1)
        summary_table_row = 2
        
           Dim row As Double
        row = 2
        Dim column As Integer
        column = 1
     
        
            ws.Cells(1, 9).Value = "Ticker_symbol"
            ws.Cells(1, 10).Value = "Volume_total"
            ws.Cells(1, 11).Value = "YearChange"
            ws.Cells(1, 12).Value = "YearPChange"
            
        Open_Price = Cells(2, column + 2).Value

        For i = 2 To 70926
            If Cells(i + 1, 1).Value <> Cells(i, 1).Value Then
            'Setting Ticker Name
                Ticker_symbol = Cells(i, 1).Value
            'Percent Change
            If (Open_Price = 0 And Close_Price = 0) Then
                YearPChange = 0
            ElseIf (Open_Price = 0 And Close_Price <> 0) Then
                YearPChange = 1
            Else
                YearPChange = Year_Change / Open_Price
                Cells(row, column + 10).Value = YearPChange
                Cells(row, column + 10).NumberFormat = "0.00%"
            End If
            
            
            
        For i = 2 To 70926
        
        If Cells(i + 1, 1).Value <> Cells(i, 1).Value Then
        
        volume_total = volume_total + Cells(i, column + 6).Value
        
            'Year Change
            Year_Change = Close_Price - Open_Price
            Cells(row, column + 9).Value = Year_Change
            'Need to print headers
               
                Range("I" & summary_table_row).Value = Ticker_symbol
                
                Range("J" & summary_table_row).Value = volume_total

                Range("K" & summary_table_row).Value = Year_Change

                Range("L" & summary_table_row).Value = YearPChange

                summary_table_row = summary_table_row + 1
                
            Else

                'volume_total = volume_total + Cells(i, column + 6).Value
            End If
            
        Next i
    Next
End Sub

' * You should also have conditional formatting that will highlight positive change in green and negative change in red.
    '
    
    ' If (Year_change > 0) then

    '     Cells(i, j).Interior.ColorIndex(green)
    ' IF (YearPChange > 0 ) then

    '     Cells(i, j).Interior.COlorIndex(green)
' * The result should look as follows.

' ![moderate_solution](Images/moderate_solution.png)

' ### CHALLENGES

' 1. Your solution will also be able to return the stock with the "Greatest % increase", "Greatest % Decrease" and "Greatest total volume". The solution will look as follows:

' ![hard_solution](Images/hard_solution.png)

' 2. Make the appropriate adjustments to your VBA script that will allow it to run on every worksheet, i.e., every year, just by running the VBA script once.

' ### Other Considerations

' * Use the sheet `alphabetical_testing.xlsx` while developing your code. This data set is smaller and will allow you to test faster. Your code should run on this file in less than 3-5 minutes.

' * Make sure that the script acts the same on each sheet. The joy of VBA is to take the tediousness out of repetitive task and run over and over again with a click of the button.



