Sub SummaryStats()

' Cast variables
Dim TickerName As String
Dim SummaryTableIndex As Double
Dim InitialValue As Double
Dim FinalValue As Double
Dim VolumeSum As Double

' Bonus table variable definitions
Dim GreatestPerInc As Double
Dim GreatestPerDec As Double
Dim GreatestVol As Double

    ' Loop over worksheets
    For Each ws In ActiveWorkbook.Worksheets

        ' define initial values for variables
        SummaryTableIndex = 2
        InitialValue = ws.Cells(2, 3).Value
        VolumeSum = 0

        GreatestPerDec = 0
        GreatestPerInc = 0
        GreatestVol = 0
    
        ' Configure table col names: basic table
        ws.Cells(1, 9).Value = "Ticker"
        ws.Cells(1, 10).Value = "Yearly Absolute Change"
        ws.Cells(1, 11).Value = "Yearly % Change"
        ws.Cells(1, 12).Value = "Tot. Volume"

        ' Bonus task table
        ws.Cells(2, 14).Value = "Greatest % Increase"
        ws.Cells(3, 14).Value = "Greatest % Decrease"
        ws.Cells(4, 14).Value = "Greatest Tot. Volume"
        ws.Cells(1, 15).Value = "Ticker"
        ws.Cells(1, 16).Value = "Value"
    
        ' Loop over rows
        For I = 2 To ws.Cells(Rows.Count, 1).End(xlUp).Row

            'Conditional sentence

            If (ws.Cells(I, 1).Value <> ws.Cells(I + 1, 1).Value) Then
                
                TickerName = ws.Cells(I, 1).Value
                FinalValue = ws.Cells(I, 6).Value
                VolumeSum = VolumeSum + ws.Cells(I, 7).Value
                
                ' Assign to summary table
                
                ws.Cells(SummaryTableIndex, 9).Value = TickerName
                ws.Cells(SummaryTableIndex, 10).Value = FinalValue - InitialValue
                
                If (FinalValue = 0 And InitialValue = 0) Then
                    ws.Cells(SummaryTableIndex, 11).Value = 0
                    ws.Cells(SummaryTableIndex, 11).NumberFormat = "0.00%"
                    ws.Cells(SummaryTableIndex, 11).Interior.ColorIndex = 2
                Else
                    ws.Cells(SummaryTableIndex, 11).Value = (FinalValue - InitialValue) / InitialValue
                    ws.Cells(SummaryTableIndex, 11).NumberFormat = "0.00%"
                    
                    'Color coding the cells
                    If ((FinalValue - InitialValue) / InitialValue > 0) Then
                        ws.Cells(SummaryTableIndex, 11).Interior.ColorIndex = 4
                    Else
                        ws.Cells(SummaryTableIndex, 11).Interior.ColorIndex = 3
                    End If

                End If

                ws.Cells(SummaryTableIndex, 12).Value = VolumeSum

                ' Bonus summary table data assignment

                'Greatest % increase
                If (ws.Cells(SummaryTableIndex, 11).Value > GreatestPerInc) Then
                    ws.Cells(2, 15).Value = ws.Cells(SummaryTableIndex, 9).Value ' greatest % increase ticker name
                    ws.Cells(2, 16).Value = ws.Cells(SummaryTableIndex, 11).Value ' greatest % increase value
                    ws.Cells(2, 16).NumberFormat = "0.00%"
                    GreatestPerInc = ws.Cells(SummaryTableIndex, 11).Value
                End If

                'Greatest % decrease
                If (ws.Cells(SummaryTableIndex, 11).Value < GreatestPerDec) Then
                    ws.Cells(3, 15).Value = ws.Cells(SummaryTableIndex, 9).Value ' greatest % dec ticker name
                    ws.Cells(3, 16).Value = ws.Cells(SummaryTableIndex, 11).Value ' greatest % dec value
                    ws.Cells(3, 16).NumberFormat = "0.00%"
                    GreatestPerDec = ws.Cells(SummaryTableIndex, 11).Value
                End If

                'Greatest volume
                If (ws.Cells(SummaryTableIndex, 12).Value > GreatestVol) Then
                    ws.Cells(4, 15).Value = ws.Cells(SummaryTableIndex, 9).Value ' greatest vol ticker name
                    ws.Cells(4, 16).Value = ws.Cells(SummaryTableIndex, 12).Value ' greatest vol value
                    GreatestVol = ws.Cells(SummaryTableIndex, 12).Value
                End If
                
                ' Rename initial value for next Ticker,  set new index & reset total
                SummaryTableIndex = SummaryTableIndex + 1
                VolumeSum = 0

                If (ws.Cells(I + 1, 3).Value = 0) Then
                    For P = 1 To 365
                        If (ws.Cells(I + P, 3).Value > 0) Then
                            InitialValue = ws.Cells(I + P, 3).Value
                            Exit For
                        End If
                    Next P
                Else
                    InitialValue = ws.Cells(I + 1, 3).Value
                End If

            Else

                VolumeSum = VolumeSum + ws.Cells(I, 7).Value
                
            End If
        
        Next I
    
    Next ws
    
    MsgBox ("Summary completed! :-D")

End Sub