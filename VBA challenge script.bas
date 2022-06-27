Attribute VB_Name = "Module1"
Sub VBAchallenge():

'First set up your loop to go through all worksheets

    Dim Worksheet As Worksheet
        
    For Each Worksheet In ActiveWorkbook.Worksheets
        
        

     'define your variables
            Dim ticker As String
            Dim totalStockVolume As Double
            totalStockVolume = 0
    
            'identify where they will be located on your excel doc
            Dim summaryTableRow As Integer
            summaryTableRow = 2
           
            Dim openPrice As Double
            openPrice = Cells(2, 3).Value
            
            Dim closePrice As Double
            
            Dim yearly_change As Double
            Dim percent_change As Double
    
            'Label the Summary Table headers
            Cells(1, 9).Value = "Ticker"
            Cells(1, 10).Value = "Yearly Change"
            Cells(1, 11).Value = "Percent Change"
            Cells(1, 12).Value = "Total Stock Volume"
    
            'Start script to loop through the ticker values
            
            'identify your last row
            LastRow = Cells(Rows.Count, 1).End(xlUp).Row
    
            For i = 2 To LastRow
    
                If Cells(i + 1, 1).Value <> Cells(i, 1).Value Then
            
                 'identify ticker
                  ticker = Cells(i, 1).Value
    
                  'identify volume
                  totalStockVolume = totalStockVolume + Cells(i, 7).Value
    
                  'Print the ticker name in the summary table
                  Range("I" & summaryTableRow).Value = ticker
    
                  'Print the trade volume for each ticker in the summary table
                  Range("L" & summaryTableRow).Value = totalStockVolume
                  
                    'Now collect information about closing price
                    closePrice = Cells(i, 6).Value

    
                  'Calculate yearly change
                  yearly_change = (closePrice - openPrice)
                  
                  'Print the yearly change for each ticker in the summary table
                  Range("J" & summaryTableRow).Value = yearly_change
    
                 'Once you have your yearly change, you can then create your IF statement to find percent
                    If (openPrice = 0) Then
    
                        percent_change = 0
    
                    Else
                        
                        percent_change = yearly_change / openPrice
                    
                    End If
    
                  'Print the yearly change for each ticker in the summary table
                  Range("K" & summaryTableRow).Value = percent_change
                  Range("K" & summaryTableRow).NumberFormat = "0.00%"
       
                  'Reset the row counter. Add one to the summaryTableRow
                  summaryTableRow = summaryTableRow + 1
    
                  'Reset totalStockVolume
                  totalStockVolume = 0
    
                  'Reset the openPrice
                  openPrice = Cells(i + 1, 3)
                
                Else
                  
                  'add the volume
                  totalStockVolume = totalStockVolume + Cells(i, 7).Value
    
                
                End If
            
            Next i
    
        'include conditional formatting so you can assign colors
    
        lastrow_summary_table = Cells(Rows.Count, 9).End(xlUp).Row
        
        For i = 2 To lastrow_summary_table
                If Cells(i, 10).Value > 0 Then
                    Cells(i, 10).Interior.ColorIndex = 10
                Else
                    Cells(i, 10).Interior.ColorIndex = 3
                End If
        Next i
    
    Next Worksheet

End Sub
