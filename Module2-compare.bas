Attribute VB_Name = "Module2"
Sub compare()
   
   For Each ws In Worksheets
   
        Dim last_row As Integer
        Dim greatestIncrease As Double
        Dim greatestDecrease As Double
        Dim greatestTotalVolume As Double
        
        Dim ticker1 As String
        Dim ticker2 As String
        Dim ticker3 As String
   
        ws.Cells(1, 17).Value = "Ticker"
        ws.Cells(1, 18).Value = "value"
        ws.Cells(2, 15).Value = "Greatest%Increase"
        ws.Cells(3, 15).Value = "Greatest%Decrease"
        ws.Cells(4, 15).Value = "GreatestTotalVolume"
        last_row = ws.Cells(Rows.Count, 10).End(xlUp).Row
        
        For i = 2 To last_row
           If (i = 2) Then
              greatestIncrease = ws.Cells(i, 12).Value
              greatestDecrease = ws.Cells(i, 12).Value
              greatestTotalVolume = ws.Cells(i, 13).Value
              ticker1 = ws.Cells(i, 10).Value
              ticker2 = ws.Cells(i, 10).Value
              ticker3 = ws.Cells(i, 10).Value
       
           End If
   
           If (ws.Cells(i, 12).Value > greatestIncrease) Then
               greatestIncrease = ws.Cells(i, 12).Value
               ticker1 = ws.Cells(i, 10).Value
       
           End If
     
           If (ws.Cells(i, 12).Value < greatestDecrease) Then
             greatestDecrease = ws.Cells(i, 12).Value
             ticker2 = ws.Cells(i, 10).Value
           End If
     
           If (ws.Cells(i, 13).Value > greatestTotalVolume) Then
             greatestTotalVolume = ws.Cells(i, 13).Value
             ticker3 = ws.Cells(i, 10).Value
          End If
     
      Next i
   
      ws.Cells(2, 18).Value = FormatPercent(greatestIncrease)
      ws.Cells(3, 18).Value = FormatPercent(greatestDecrease)
      ws.Cells(4, 18).Value = greatestTotalVolume
      ws.Cells(2, 17).Value = ticker1
      ws.Cells(3, 17).Value = ticker2
      ws.Cells(4, 17).Value = ticker3
      greatestTotalVolume = 0
    Next ws
      
   
   
End Sub
