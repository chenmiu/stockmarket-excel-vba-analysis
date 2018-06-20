Sub moderate_script()

    Dim imax As Long
    Dim volsum As Double
    Dim isym As Integer
    Dim symstart As Double
    Dim symend As Double
    

    Dim Current As Worksheet
    
    For Each ws In Worksheets
    
        ws.Activate
    
        imax = Cells(Rows.Count, 1).End(xlUp).Row
    
        isym = 2
        volsum = 0
    
        Cells(1, 9).Value = "Ticker"
        Cells(1, 10).Value = "Yearly Change"
        Cells(1, 11).Value = "Percent Change"
        Cells(1, 12).Value = "Total Volume"
    
        symstart = Cells(2, 3).Value
    
        For i = 2 To imax
            volsum = volsum + Cells(i, 7).Value
            If Cells(i, 1).Value <> Cells(i + 1, 1).Value Then
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
                volsum = 0
                symstart = Cells(i + 1, 3).Value
                isym = isym + 1
            End If
        Next i
        
    Next ws

End Sub
