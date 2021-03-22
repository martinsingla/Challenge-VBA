Sub SummaryStats()

' Cast variables
Dim TickerName As String
Dim SummaryTableIndex As Double
Dim InitialValue As Double
Dim FinalValue As Double
Dim VolumeSum As Double

    ' Loop over worksheets
    For Each ws in ActiveWorkbook.Worksheets

        ' define initial values for variables
        SummaryTableIndex = 2
        InitialValue = ws.Cells(2, 3).Value
        VolumeSum = 0
    
        ' Configure table col names
        ws.Cells(1, 9).Value = "Ticker"
        ws.Cells(1, 10).Value = "Yearly Absolute Change"
        ws.Cells(1, 11).Value = "Yearly % Change"
        ws.Cells(1, 12).Value = "Tot. Volume"
    
        ' Loop over rows
        For i = 2 To ws.Cells(Rows.Count, 1).End(xlUp).Row

            'Conditional sentence

            If (ws.Cells(i, 1).Value <> ws.Cells(i + 1, 1).Value) Then
                
                TickerName = ws.Cells(i, 1).Value
                FinalValue = ws.Cells(i, 6).Value
                VolumeSum = VolumeSum + ws.Cells(i, 7).Value
                
                ' Assign to summary table
                
                ws.Cells(SummaryTableIndex, 9).Value = TickerName
                ws.Cells(SummaryTableIndex, 10).Value = FinalValue - InitialValue
                
                If (FinalValue = 0 And InitialValue = 0) Then
                    ws.Cells(SummaryTableIndex, 11).Value = 0
                    ws.Cells(SummaryTableIndex, 11).NumberFormat="0.00%"
                    ws.Cells(SummaryTableIndex, 11).Interior.ColorIndex= 2
                Else
                    ws.Cells(SummaryTableIndex, 11).Value = (FinalValue - InitialValue) / InitialValue
                    ws.Cells(SummaryTableIndex, 11).NumberFormat="0.00%"
                    
                    If ((FinalValue - InitialValue) / InitialValue > 0) Then
                        ws.Cells(SummaryTableIndex, 11).Interior.ColorIndex= 4
                    Else 
                        ws.Cells(SummaryTableIndex, 11).Interior.ColorIndex= 3
                    End If

                End If

                ws.Cells(SummaryTableIndex, 12).Value = VolumeSum
                
                ' Rename initial value for next Ticker,  set new index & reset total
                SummaryTableIndex = SummaryTableIndex + 1
                InitialValue = ws.Cells(i+1, 3).Value
                VolumeSum = 0
            
            Else

                VolumeSum = VolumeSum + ws.Cells(i, 7).Value
                
            End If
        
        Next I
    
    Next ws
    
    MsgBox("Summary completed! :-D")

End Sub

