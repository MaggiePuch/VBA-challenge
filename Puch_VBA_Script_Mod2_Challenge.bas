Attribute VB_Name = "Module1"
Sub Ticker()

  For Each w In Worksheets
    'Initial variables for Ticker Macro
    Dim Ticker As String
    Dim NextTicker As String
    Dim PreTicker As String
    Dim Volume As Double
    Dim Yearly As Double
    Dim Opening As Double
    Dim Closing As Double
    Dim lastrow As Double
    Dim Summary_Table_Row As Integer
    Dim Percent As Double
    
    'Initial Variable for Greatest

    
    'Initial values set to 0 for variables
    Volume = 0
    Yearly = 0
    Opening = 0
    Closing = 0
    Summary_Table_Row = 2
    Increase = 0
    Decrease = 0
    Percent = 0

    
    'Set last row
    lastrow = w.Cells(Rows.Count, 1).End(xlUp).Row
    
    'Set titles for summary table
        w.Range("I1").Value = "Ticker"
        w.Range("J1").Value = "Yearly Change"
        w.Range("K1").Value = "Percent Change"
        w.Range("L1").Value = "Total Stock Volume"
        w.Range("P1").Value = "Ticker"
        w.Range("Q1").Value = "Value"
        w.Range("O2").Value = "Greatest % Increase"
        w.Range("O3").Value = "Greatest % Decrease"
        w.Range("O4").Value = "Greatest Total Volume"
        
 'Loop through all ticker rows
    For i = 2 To lastrow
          'Set Values
            Ticker = w.Cells(i, 1).Value
            NextTicker = w.Cells(i + 1, 1).Value
            PreTicker = w.Cells(i - 1, 1).Value
            Volume = Volume + w.Cells(i, 7).Value
        'Check if we are still in the same ticker, if it is not...
    
    If NextTicker <> Ticker Then
          'Closing Price
            Closing = w.Cells(i, 6).Value
          'Add to Yearly
            Yearly = Closing - Opening
          'Add Percent Change
            Percent = (Closing - Opening) / Opening
          'Print Ticker in Summary Table
            w.Range("I" & Summary_Table_Row).Value = Ticker
          'Print Volume to Summary Table
            w.Range("L" & Summary_Table_Row).Value = Volume
           'Print Yearly in Summary Table
            w.Range("J" & Summary_Table_Row).Value = Yearly
           'Print Percent Change and format to Percent
            w.Range("K" & Summary_Table_Row).Value = Percent
            w.Range("K" & Summary_Table_Row).NumberFormat = "0.00%"
            'Color for Yearly Change
         If Yearly >= 0 Then
                w.Range("J" & Summary_Table_Row).Interior.ColorIndex = 4
            Else: w.Range("J" & Summary_Table_Row).Interior.ColorIndex = 3
        
          End If
            
        'Add one to the summary table row
            Summary_Table_Row = Summary_Table_Row + 1
         'Reset the Volume Total
            Volume = 0
         'Reset the Yearly Total
            Yearly = 0
            
          'If the cell immediatly following a row is the same ticker...
         ElseIf PreTicker <> Ticker Then
            Opening = w.Cells(i, 3).Value
            
      End If
    Next i
'Max,Min,Greatest Volume
    w.Range("Q2").Value = Application.WorksheetFunction.Max(w.Range("K2:K" & lastrow))
    w.Range("Q2:Q3").NumberFormat = "0.00%"
    w.Range("Q3").Value = Application.WorksheetFunction.Min(w.Range("K2:K" & lastrow))
    w.Range("Q4").Value = Application.WorksheetFunction.Max(w.Range("L2:L" & lastrow))

'Match for Ticker of Max,Min,Greatest Volume
    w.Range("P2").Value = Application.WorksheetFunction.Index(w.Range("I2:I" & lastrow), Application.WorksheetFunction.Match(w.Range("Q2").Value, w.Range("K2:K" & lastrow), 0))
    w.Range("P3").Value = Application.WorksheetFunction.Index(w.Range("I2:I" & lastrow), Application.WorksheetFunction.Match(w.Range("Q3").Value, w.Range("K2:K" & lastrow), 0))
    w.Range("P4").Value = Application.WorksheetFunction.Index(w.Range("I2:I" & lastrow), Application.WorksheetFunction.Match(w.Range("Q4").Value, w.Range("L2:L" & lastrow), 0))

'Format width of columns
    w.Range("A:Q").EntireColumn.AutoFit

Next w

End Sub

