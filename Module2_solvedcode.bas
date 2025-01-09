Attribute VB_Name = "Module1"
Sub StockAnalysis()

'Define all variables
    Dim ws As Worksheet
    Dim ticker As String
    Dim year_open As Double
    Dim year_close As Double
    Dim yearly_change As Double
    Dim percent_change As Double
    Dim total_volume As LongLong
        Dim max_increase As Double
        Dim max_increase_ticker As String
        Dim max_decrease As Double
        Dim max_decrease_ticker As String
        Dim max_volume As LongLong
        Dim max_volume_ticker As String
    Dim Summary_Table_Row As Integer
    Dim lastrow As Long
    
'Apply code to all worksheets
    For Each ws In ThisWorkbook.Worksheets
    
'Assign starting values to variables
    Summary_Table_Row = 2
    total_volume = 0
    year_open = ws.Cells(2, 3).Value
    year_close = 0
    lastrow = ws.Cells(Rows.Count, 1).End(xlUp).Row
    max_increase = 0
    max_decrease = 0
    max_volume = 0
    
'Add headers and data to summary table

        ws.Range("I1").Value = "Ticker"
        ws.Range("J1").Value = "Yearly Change"
        ws.Range("K1").Value = "Percent Change"
        ws.Range("L1").Value = "Total Stock Volume"
    'Bonus table headings
        ws.Range("O2").Value = "Greatest % Increase"
        ws.Range("O3").Value = "Greatest % Decrease"
        ws.Range("O4").Value = "Greatest Total Volume"
        ws.Range("P1").Value = "Ticker"
        ws.Range("Q1").Value = "Value"

'begin looping through sheets
    
    For i = 2 To lastrow
            
          If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then
          
          'Name ticker
          ticker = ws.Cells(i, 1).Value
          
          'Add final line to total stock volume
          total_volume = total_volume + ws.Cells(i, 7).Value
    
          'Set year close value
          year_close = ws.Cells(i, 6).Value
          
          'Print ticker to summary table
          ws.Range("I" & Summary_Table_Row).Value = ticker
          
          'Print yearly change to summary table
          ws.Range("J" & Summary_Table_Row).Value = year_close - year_open
                'Add conditional formating to summary table
            If ws.Range("J" & Summary_Table_Row).Value > 0 Then
                ws.Range("J" & Summary_Table_Row).Interior.ColorIndex = 4
            ElseIf ws.Range("J" & Summary_Table_Row).Value <= 0 Then
                ws.Range("J" & Summary_Table_Row).Interior.ColorIndex = 3
          
          End If
          
          'Add value to percent_change
          percent_change = (year_close - year_open) / year_open
          'Print percent change to summary table
            ws.Range("K" & Summary_Table_Row).Value = percent_change
             'Format percent changed
                  ws.Range("K" & Summary_Table_Row).NumberFormat = "0.00%"
          
          'Print total volume to summary table
          ws.Range("L" & Summary_Table_Row).Value = total_volume
          
                  'Max and Min calculations
            If percent_change > max_increase Then
                max_increase = percent_change
                max_increase_ticker = ticker
                
            ElseIf percent_change < max_decrease Then
                max_decrease = percent_change
                max_decrease_ticker = ticker
            End If

        If total_volume > max_volume Then
            max_volume = total_volume
            max_volume_ticker = ticker
            End If
            
        'print to table
         ws.Cells(4, 16).Value = max_volume_ticker
         ws.Cells(4, 17).Value = max_volume
         ws.Cells(2, 16).Value = max_increase_ticker
         ws.Cells(2, 17).Value = max_increase
         ws.Cells(3, 16).Value = max_decrease_ticker
         ws.Cells(3, 17).Value = max_decrease
         
         'format table
         ws.Cells(2, 17).NumberFormat = "0.00%"
         ws.Cells(3, 17).NumberFormat = "0.00%"
          
          'Set variables for next loop
            Summary_Table_Row = Summary_Table_Row + 1
            total_volume = 0
            year_open = ws.Cells(i + 1, 3).Value
            year_close = 0

         
    End If
  
        'Add up all volume
        total_volume = total_volume + ws.Cells(i, 7).Value
         
    Next i
    
    'Format summary table columns
        ws.Range("I1").EntireColumn.AutoFit
        ws.Range("J1").EntireColumn.AutoFit
        ws.Range("K1").EntireColumn.AutoFit
        ws.Range("L1").EntireColumn.AutoFit
        ws.Range("O1").EntireColumn.AutoFit
        
        
        
        
    Next ws

  



End Sub

