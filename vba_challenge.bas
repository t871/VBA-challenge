Attribute VB_Name = "Module2"
Sub stock_market_analysis()
    For Each ws In Worksheets
        Dim ticker_name As String
        Dim last_row As Long
        Dim total_volume As Double
        total_volume = 0
        Dim table As Long
        table = 2
        Dim year_open As Double
        Dim year_close As Double
        Dim year_change As Double
        Dim previous As Long
        previous = 2
        Dim percent_change As Double
             
        last_row = ws.Cells(Rows.Count, 1).End(xlUp).Row
        
        For i = 2 To last_row
            total_volume = total_volume + ws.Cells(i, 7).Value
           
            If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then
                ticker_name = ws.Cells(i, 1).Value
                ws.Range("I" & table).Value = ticker_name
                ws.Range("L" & table).Value = total_volume
                total_volume = 0
                year_open = ws.Range("C" & previous)
                year_close = ws.Range("F" & i)
                year_change = year_close - year_open
                ws.Range("J" & table).Value = year_change

            
                If year_open = 0 Then
                    percent_change = 0
                Else
                    year_open = ws.Range("C" & previous)
                    percent_change = year_change / year_open
                End If
           
                ws.Range("K" & table).NumberFormat = "0.00%"
                ws.Range("K" & table).Value = percent_change

               
                If ws.Range("J" & table).Value >= 0 Then
                    ws.Range("J" & table).Interior.ColorIndex = 4
                Else
                    ws.Range("J" & table).Interior.ColorIndex = 3
                End If
            
              
                table = table + 1
                previous = i + 1
                End If
            Next i
            
            Next ws
            

End Sub

