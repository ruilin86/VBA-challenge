Sub stock_analysis()
 
For Each ws In Worksheets

'define variables
  
    Dim Ticker As String
    Dim total As Double
    Dim OpenPrice As Double
    Dim ClosePrice As Double
    Dim rowCount As Double
    Dim PriceChange As Double
    Dim PercentChange As Double
    Dim i As Double
    Dim j As Double 'store the row# for open price
    Dim a As Double 'store the row # for print location
    Dim b As Double 'store the row# of first non-zero open price for each ticker
  

    'set range for output
     ws.Range("I1").Value = "Ticker"
     ws.Range("J1").Value = "Yearly Change"
     ws.Range("K1").Value = "Percent Change"
     ws.Range("L1").Value = "Total Stock Volume"
     ws.Range("O2").Value = "Greatest % Increase"
     ws.Range("O3").Value = "Greatest % Decrease"
     ws.Range("O4").Value = "Greatest Total Volumn"
     ws.Range("P1").Value = "Ticker"
     ws.Range("Q1").Value = "Value"
  
    'count row number
     rowCount = ws.Cells(Rows.Count, "A").End(xlUp).Row
     
    'set initial value
     total = 0
     j = 2
     a = 2
     b = 2
     
 
 'loop to find the open price and close price for each ticker
  For i = 2 To rowCount
      
      If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then
         
          total = total + ws.Cells(i, 7).Value
          
          'handle total 0 volumn
          If total = 0 Then
          
             ws.Range("I" & a).Value = ws.Cells(i, 1).Value
             ws.Range("J" & a).Value = 0
             ws.Range("K" & a).Value = "%" & 0
             ws.Range("L" & a).Value = 0
          
           'if the first open price is 0, find the first non-zero open price
          Else
           
          'if the fisrt open price is 0, find the next non-zero open price
          
           If ws.Cells(j, 3).Value = 0 Then
              For b = j To i
                If ws.Cells(b, 3).Value <> 0 Then
                   j = b
                   Exit For
                End If
               Next b
            End If
        
          
          Ticker = ws.Cells(i, 1).Value
          ClosePrice = ws.Cells(i, 6).Value
          OpenPrice = ws.Cells(j, 3).Value
          PriceChange = ClosePrice - OpenPrice
          PercentChange = PriceChange / OpenPrice
          
          
  
        'print results
         ws.Range("I" & a).Value = Ticker
         ws.Range("J" & a).Value = PriceChange
         ws.Range("K" & a).Value = PercentChange
         ws.Range("K" & a).NumberFormat = "0.00%"
         ws.Range("L" & a).Value = total
         
         
         'format stock price change
         If PriceChange < 0 Then
            ws.Range("J" & a).Interior.Color = vbRed
       
         ElseIf PriceChange > 0 Then
              ws.Range("J" & a).Interior.Color = vbGreen
         
          Else: ws.Range("J" & a).Interior.ColorIndex = 0
             
         End If
     End If
    
    j = i + 1
    total = 0
    a = a + 1
   
         
   Else
       total = total + ws.Cells(i, 7).Value
      
  End If
         
   Next i
   
   ws.Range("Q2") = WorksheetFunction.Max(ws.Range("K:K"))
   ws.Range("Q3") = WorksheetFunction.Min(ws.Range("K:K"))
   ws.Range("Q4") = WorksheetFunction.Max(ws.Range("L:L"))
   ws.Range("Q2").NumberFormat = "0.00%"
   ws.Range("Q3").NumberFormat = "0.00%"
   
   Dim increase_num, decrease_num, volume_num As Double
   increase_num = WorksheetFunction.Match(ws.Range("Q2"), ws.Range("K:K"), 0)
   decrease_num = WorksheetFunction.Match(ws.Range("Q3"), ws.Range("K:K"), 0)
   volume_num = WorksheetFunction.Match(ws.Range("Q4"), ws.Range("L:L"), 0)
   
   ws.Range("P2") = ws.Range("I" & increase_num)
   ws.Range("P3") = ws.Range("I" & decrease_num)
   ws.Range("P4") = ws.Range("I" & volume_num)
  
   
Next ws

         
End Sub