Attribute VB_Name = "Module1"
Sub Stock_Data()
    
  Dim Ticker As String
  Dim Yearly_Change As Double
  Dim Percent_Change As Variant
  Dim Total_Volume As LongLong
  Dim ws As Worksheet
  For Each ws In Worksheets
  Yearly_Change = 0
  Percent_Change = 0
  Total_Volume = 0
      
      
  ws.Cells(1, 9) = "Ticker"
  ws.Cells(1, 10) = "Yearly Change"
  ws.Cells(1, 11) = "Percent Change"
  ws.Cells(1, 12) = "Total Volume"
  ws.Cells(1, 16) = "Ticker"
  ws.Cells(1, 17) = "Value"
  ws.Cells(2, 15) = "Greatest % Increase"
  ws.Cells(3, 15) = "Greatest % Decrease"
  ws.Cells(4, 15) = "Greatest Total Volume"
  
  ws.Columns("O").ColumnWidth = 20
    
  LastRow = ws.Cells(Rows.Count, 1).End(xlUp).Row
    
  Open_Value = ws.Cells(2, 3).Value
    Start_Row = 2
  
  Dim Summary_Table_Row As Integer
  Summary_Table_Row = 2

  For i = 2 To LastRow

    If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then
        ' Add Total_Volume
     Total_Volume = Total_Volume + ws.Cells(i, 7).Value
      Ticker = ws.Cells(i, 1).Value
      End_Row = i
      If Total_Volume = 0 Then
        Yearly_Change = 0
        Percent_Change = 0
    
      Else
        If Open_Value = 0 Then
            For j = Start_Row To End_Row
                If Cells(j, 3).Value <> 0 Then
                    Open_Value = Cells(j, 3).Value
                    Exit For
                End If
            Next j
        End If
        
        Yearly_Change = ws.Cells(i, 6).Value - Open_Value
      
        Percent_Change = Yearly_Change / Open_Value
        
        If ws.Range("K" & Summary_Table_Row) >= 0 Then
            ws.Range("K" & Summary_Table_Row).Interior.ColorIndex = 4
        Else: ws.Range("K" & Summary_Table_Row).Interior.ColorIndex = 3
        
        End If
        
      End If
      
      
      ' Add Total_Volume
     Total_Volume = Total_Volume + ws.Cells(i, 7).Value
       
        
      ' Print the Ticker in the Summary Table
      ws.Range("I" & Summary_Table_Row).Value = Ticker
      
      'Print the Yearly Change in the Summary Table
      ws.Range("J" & Summary_Table_Row).Value = Yearly_Change
      
      'Print Percent Change in the Summary Table
      ws.Range("K" & Summary_Table_Row).NumberFormat = "0.00%"
      ws.Range("K" & Summary_Table_Row).Value = Percent_Change
     
      ' Print the Total Volume to the Summary Table
      ws.Range("L" & Summary_Table_Row).Value = Total_Volume

      ' Add one to the summary table row
      Summary_Table_Row = Summary_Table_Row + 1
      
      Open_Value = ws.Cells(i + 1, 3).Value
      Start_Row = i + 1
      
      If ws.Range("K" & Summary_Table_Row) >= 0 Then
            ws.Range("K" & Summary_Table_Row).Interior.ColorIndex = 4
        Else: ws.Range("K" & Summary_Table_Row).Interior.ColorIndex = 3
        
        End If
      
      
      ' Reset the Table Total
      Total_Volume = 0
      Yearly_Change = 0
      Percent_Change = 0

    ' If the cell immediately following a row is the same Ticker...
    Else

      ' Add to the Total_Volume
      Total_Volume = Total_Volume + ws.Cells(i, 7).Value
      
    
    End If
    
    Next i

    Next ws
    
End Sub



