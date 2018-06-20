Sub collectdata()

    Dim imax As Long
    Dim volsum As Double
    Dim isym As Integer

    Dim Current As Worksheet
    
    For Each ws In Worksheets
    
        ws.Activate
    
        imax = Cells(Rows.Count, 1).End(xlUp).Row
    
        isym = 2
        volsum = 0
    
        Cells(1, 9).Value = "Ticker"
        Cells(1, 10).Value = "Total Volume"
    
        For i = 2 To imax
            volsum = volsum + Cells(i, 7).Value
            If Cells(i, 1).Value <> Cells(i + 1, 1).Value Then
                Cells(isym, 9).Value = Cells(i, 1).Value
                Cells(isym, 10).Value = volsum
                volsum = 0
                isym = isym + 1
            End If
        Next i
    Next ws

End Sub
