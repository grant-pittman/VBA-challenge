Sub stock_tracker()
    'Set a variable for worksheets
    Dim ws As Worksheet

  ' Set an initial variable for holding the ticker symbol
  Dim ticker As String

  ' Set an initial variable for holding the total traded volume and initializing at 0
  Dim total_volume As Double
  total_volume = 0

  ' Keep track of the location for each ticker in the summary table
  Dim Summary_Table_Row As Integer
  Summary_Table_Row = 2
  
  'keeping track of how many days of stock data. will use to calculate yearly changes
  Dim days As Integer
  days = 0
  
  'set a variable for the yearly change of an indivudual stock
  Dim yearly_change As Double
  
  'set a variable for the percent change of an individual stock
  Dim percent_change As Double
  
  Dim greatest_increase As Double
  Dim greatest_decrease As Double
  Dim greatest_volume As Double
 ' greatest_volume = ws.Range("L2").Value
  
  

  'loop through all worksheets

For Each ws In Worksheets
    'setting initial value for the last row of stock data to be analysed
    LastRow = ws.Cells(Rows.Count, 1).End(xlUp).Row
    
    'Creating column headers for summary table
    ws.Cells(1, 9).Value = "Ticker"
    ws.Cells(1, 10).Value = "Yearly Change"
    ws.Cells(1, 11).Value = "Percent Change"
    ws.Cells(1, 12).Value = "Total Stock Volume"
    ws.Cells(1, 13).Value = "Greatest Percent Increase"
    ws.Cells(1, 14).Value = "Greatest Percent Decrease"
    ws.Cells(1, 15).Value = "Greatest Total Volume"
    
    'looping through the stock data
  For I = 2 To LastRow

    ' Check if we are still within the same stock ticker, if it is not...
    If ws.Cells(I + 1, 1).Value <> ws.Cells(I, 1).Value Then

      ' Set the ticker name
      ticker = ws.Cells(I, 1).Value

        'sets the yearly change of the stock
      yearly_change = ws.Cells(I, 6).Value - ws.Cells(I - days, 3).Value
      
      'below conditional solves for the divide by 0 problem
      If ws.Cells(I - days, 3).Value = 0 Then
        percent_change = 0
      Else
        percent_change = (ws.Cells(I, 6).Value - ws.Cells(I - days, 3).Value) / ws.Cells(I - days, 3).Value
      End If
        
      ' Print the ticker in the Summary Table
      ws.Range("I" & Summary_Table_Row).Value = ticker

      ' Print the total volume to the Summary Table
      ws.Range("L" & Summary_Table_Row).Value = total_volume
      
      'Print yearly change to the Summary Table
      ws.Range("J" & Summary_Table_Row).Value = yearly_change
      
            'below conditional color formats the cells. Green for positive yearly change and red for negative
      If ws.Range("J" & Summary_Table_Row).Value > 0 Then
        ws.Range("J" & Summary_Table_Row).Interior.Color = vbGreen
        
        Else
        ws.Range("J" & Summary_Table_Row).Interior.Color = vbRed
        
        End If
        
        'Print the percent change to the summary table
      ws.Range("K" & Summary_Table_Row).Value = percent_change
      
          'below conditional color formats the cells. Green for positive yearly change and red for negative
      If ws.Range("K" & Summary_Table_Row).Value > 0 Then
        ws.Range("K" & Summary_Table_Row).Interior.Color = vbGreen
        
        Else
        ws.Range("K" & Summary_Table_Row).Interior.Color = vbRed
        
        End If
      

      ' Add one to the summary table row
      Summary_Table_Row = Summary_Table_Row + 1
      
      ' Reset the total volume
      total_volume = 0
      days = 0
      
    Else

      ' Add to the volume Total
      total_volume = total_volume + ws.Cells(I, 7).Value
      'add one day
      days = days + 1
      
      
    End If
    Next I
    
    'tried to get the greatest percent increase, but ran out of time :(
    'For I = 2 To 10000
    'greatest_increase = 0
    '   If ws.Cells(I, 11).Value > greatest_increase Then
    '       greatest_increase = ws.Cells(I, 11).Value
            
    '   End If
    'Next I
    
    'ws.Range("M2").Value = greatest_increase
    
    'reset summary table row
    
    Summary_Table_Row = 2
    
    Next ws
    
End Sub



