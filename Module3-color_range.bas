Attribute VB_Name = "Module3"
Sub color_range()
   For Each ws In Worksheets
      Dim last_row As Integer
      Dim i As Integer
      last_row = ws.Cells(Rows.Count, 11).End(xlUp).Row
      For i = 2 To last_row
          If (ws.Cells(i, 11).Value > 0) Then
             ws.Cells(i, 11).Interior.ColorIndex = 4
             
          End If
          
          If (ws.Cells(i, 11).Value < 0) Then
             ws.Cells(i, 11).Interior.ColorIndex = 3
           End If
           
             
      
      Next i
      
   
   
   Next ws






End Sub
