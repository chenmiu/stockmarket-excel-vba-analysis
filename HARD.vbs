Sub hard()

    Dim imax As Long
    Dim volsum As Double
    Dim isym As Integer
    Dim symstart As Double
    Dim symend As Double
    Dim max_increase As Double
    Dim max_decrease As Double
    Dim max_volume As Double
    Dim max_increase_ticker As String
    Dim max_decrease_ticker As String
    Dim max_volume_ticker As String
    Dim symbol As String

    Dim ws As Worksheet
    
    For Each ws In Worksheets
    
        ws.Activate
    
        imax = Cells(Rows.Count, 1).End(xlUp).Row
    
        isym = 2
        volsum = 0
        max_increase = -100
        max_decrease = 100
        max_volume = -1
        symstart = Cells(2, 3).Value
        symbol = Cells(2, 1).Value

        ' set some col headers
        Cells(1, 9).Value = "Ticker"
        Cells(1, 10).Value = "Yearly Change"
        Cells(1, 11).Value = "Percent Change"
        Cells(1, 12).Value = "Total Volume"
        
        For i = 2 To imax
            volsum = volsum + Cells(i, 7).Value
            If symbol <> Cells(i + 1, 1).Value Then
                symend = Cells(i, 6).Value
                Cells(isym, 9).Value = Cells(i, 1).Value
                Cells(isym, 10).Value = symend - symstart
                If Cells(isym, 10).Value >= 0 Then
                    Cells(isym, 10).Interior.Color = RGB(0, 255, 0)
                Else
                      Cells(isym, 10).Interior.Color = RGB(255, 0, 0)
                End If
                Cells(isym, 11).Value = (symend - symstart) / (symstart + 1E-10)
                Cells(isym, 11).NumberFormat = "0.00%"
                Cells(isym, 12).Value = volsum
                
                ' for yearly summary:
                If volsum > max_volume Then
                    max_volume = volsum
                    max_volume_ticker = Cells(isym, 9).Value
                End If
                If Cells(isym, 11).Value < max_decrease Then
                    max_decrease = Cells(isym, 11).Value
                    max_decrease_ticker = Cells(isym, 9).Value
                End If
                If Cells(isym, 11).Value > max_increase Then
                    max_increase = Cells(isym, 11).Value
                    max_increase_ticker = Cells(isym, 9).Value
                End If

                'reset for next ticker
                'special handling for the case that the price is zero
                'that will force a reset on the next row
                symstart = Cells(i + 1, 3).Value
                symbol = Cells(i + 1, 1).Value
                If symstart < 0.001 Then
                    symbol = "Bogus"
                End If
                volsum = 0
                isym = isym + 1
            End If
        Next i
        
        ' output yearly summary
        Cells(1, 16).Value = "Ticker"
        Cells(1, 17).Value = "Value"
        
        Cells(2, 15).Value = "Greatest % Increase"
        Cells(2, 16).Value = max_increase_ticker
        Cells(2, 17).Value = max_increase
        Cells(2, 17).NumberFormat = "0.00%"
        
        Cells(3, 15).Value = "Greatest % Decrease"
        Cells(3, 16).Value = max_decrease_ticker
        Cells(3, 17).Value = max_decrease
        Cells(3, 17).NumberFormat = "0.00%"
        
        Cells(4, 15).Value = "Greatest Total Volume"
        Cells(4, 16).Value = max_volume_ticker
        Cells(4, 17).Value = max_volume
        
        ' scale column width appropriately
        Columns("I:Q").AutoFit
        
    Next ws

End Sub
