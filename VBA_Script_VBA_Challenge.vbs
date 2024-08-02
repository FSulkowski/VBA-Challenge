Sub Stock_Data()

Dim Ticker As String
Dim Ticker_Total_Volume As Double
Dim Quarterly_Change As Double
Dim Percent_Change As Double
Dim Summary_Table_Row As Integer
Dim lastRow As Long
Dim ws As Worksheet
Dim Previous As Double
Dim Current As Double
Dim highValue As Double
Dim lowValue As Double
Dim highVolume As Double
Dim rng1 As Range
Dim rng2 As Range
Dim rng3 As Range
Dim cell As Range


For Each ws In Worksheets

    Dim WorksheetName As String
    WorksheetName = ws.Name

    Summary_Table_Row = 2
    lastRow = Cells(Rows.Count, 1).End(xlUp).Row
    
    ws.Range("I1").Value = "Ticker"
    ws.Range("J1").Value = "Quarterly Change"
    ws.Range("K1").Value = "Percent Change"
    ws.Range("L1").Value = "Total Stock Volume"
    ws.Range("N2").Value = "Greatest % Increase"
    ws.Range("N3").Value = "Greatest % Decrease"
    ws.Range("N4").Value = "Greatest Total Volume"
    ws.Range("O1").Value = "Ticker"
    ws.Range("P1").Value = "Value"
    
'Range for Conditional Formatting
    Set rng3 = ws.Range("J2:J1501")


'Initialize for loop
For i = 2 To lastRow

'Setting range to find the highest and lowest value stocks
    Set rng1 = ws.Range("K:K")
    Set rng2 = ws.Range("L:L")
    
    'If Statement
    If Cells(i + 1, 1).Value <> Cells(i, 1).Value Then
    
    ' Open Price
        Previous = ws.Cells(i - 61, 3).Value
        
    'Closing Price
        Current = ws.Cells(i, 6).Value
        
    'Accessing Ticker Row (column A)
        Ticker = ws.Cells(i, 1).Value
    
    'Getting total volume for column L
        Ticker_Total_Volume = Ticker_Total_Volume + ws.Cells(i, 7).Value
         
   'Calculating Quarterly Change Amount
        Quarterly_Change = Current - Previous
        
    'Calculating Percentage Change
        Percent_Change = (Current / Previous) - 1
        
    'Placing Ticker Value in Column I
        ws.Range("I" & Summary_Table_Row).Value = Ticker
        
    'Placing Quarterly Change Value in Column J
        ws.Range("J" & Summary_Table_Row).Value = Quarterly_Change
        
    'Placing Percent Change Value in Column K
        ws.Range("K" & Summary_Table_Row).Value = Percent_Change

    'Placing Ticker Total Volume in Value in Column L
        ws.Range("L" & Summary_Table_Row).Value = Ticker_Total_Volume
        
    'Calculating Highest Percent from Column J and Placing it in Cells (2,16)
        highValue = WorksheetFunction.Max(rng1)
                 ws.Cells(2, 16).Value = highValue
                 ws.Cells(2, 15).Value = Ticker
                 
    'Calculating Lowest Percent from Column J and Placing it in Cells (3,16)
          lowValue = WorksheetFunction.Min(rng1)
                ws.Cells(3, 16).Value = lowValue
                ws.Cells(3, 15).Value = Ticker
                
                 
    'Retrieving the ticker for the highest and lowest values from Column A and Placing it in Cells (2,16)
        highVolume = WorksheetFunction.Max(rng2)
                 ws.Cells(4, 16).Value = highVolume
                 ws.Cells(4, 15).Value = Ticker
                    
        Summary_Table_Row = Summary_Table_Row + 1

        Ticker_Total_Volume = 0
    
        Else
  
        Ticker_Total_Volume = Ticker_Total_Volume + ws.Cells(i, 7).Value
    
    End If
    
Next i
   
   'Conditional Formatting
   
For Each cell In rng3
   
    If Not IsEmpty(cell.Value) And IsNumeric(cell.Value) Then
        If cell.Value > 0.001 Then
        cell.Interior.Color = RGB(0, 255, 0)
        
        ElseIf cell.Value < -0.001 Then
        cell.Interior.Color = RGB(255, 0, 0)
        
        Else
        cell.Interior.Color = RGB(255, 255, 255)
        
        End If
        
    End If
        
Next cell
         
    'Conditional Formatting
        rng1.NumberFormat = "0.00%"
        ws.Range("P2").NumberFormat = "0.00%"
        ws.Range("P3").NumberFormat = "0.00%"
        
 Next ws
    
End Sub