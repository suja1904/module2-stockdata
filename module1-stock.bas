Attribute VB_Name = "Module1"
Sub stock()
   
    Dim ticker As String
    Dim tickerIndex As Integer
    Dim changeIndex As Integer
   
    Dim openPrice As Double
    Dim closePrice As Double
   
    Dim volume As LongLong
    Dim last_row As LongLong
    
    
    
    For Each ws In Worksheets
      ws.Cells(1, 10).Value = "Ticker"
      ws.Cells(1, 11).Value = "Yearly Change"
      ws.Cells(1, 12).Value = "Percent Change"
      ws.Cells(1, 13).Value = "Total Stock Volume"
    
    
      last_row = ws.Cells(Rows.Count, 1).End(xlUp).Row
        
      ticker = ""
      tickerIndex = 2
      changeIndex = 2
      openPrice = 0
      closePrice = 0
      volume = 0
   
      For i = 2 To last_row
           volume = volume + ws.Cells(i, 7).Value
          'let's check the cell value with the ticker variable
           If (ws.Cells(i, 1).Value <> ticker) Then
          'ticker did not match we have found a new ticker
          'the fact is with the new ticker,the date is always the begining of the year
          'let's find out the opening price
              If (ws.Cells(i, 2).Value = "20180102" Or ws.Cells(i, 2).Value = "20190102" Or ws.Cells(i, 2).Value = "20200102") Then
             'read the opening price
                openPrice = ws.Cells(i, 3).Value
              End If
             'reading a new ticker value
             'we have found a new ticker at this point
              ticker = ws.Cells(i, 1).Value
             'we are printing the new ticker value
           
              ws.Cells(tickerIndex, 10).Value = ticker
              tickerIndex = tickerIndex + 1
           End If
           'let's check the closing date
           If (ws.Cells(i, 2).Value = "20181231" Or ws.Cells(i, 2).Value = "20191231" Or ws.Cells(i, 2).Value = "20201231") Then
              'read the closing price
               closePrice = ws.Cells(i, 6).Value
           End If
           If (openPrice <> 0 And closePrice <> 0) Then
               ws.Cells(changeIndex, 11).Value = closePrice - openPrice
              'let's calculate the percentage change
              'percentage change = yearlychange/openingPrice*100
               ws.Cells(changeIndex, 12).Value = FormatPercent(((closePrice - openPrice) / openPrice), 2)
               ws.Cells(changeIndex, 13).Value = volume
           
              changeIndex = changeIndex + 1
              openPrice = 0
              closePrice = 0
              volume = 0
          End If
     Next i
   Next ws
   
End Sub

