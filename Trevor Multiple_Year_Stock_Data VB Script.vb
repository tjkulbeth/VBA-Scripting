Sub stockinfo()
    Dim lastrow As Long
    Dim ws As Worksheet
    Dim current_ticker As String
    Dim next_ticker As String
    Dim Y_change As Variant
    Dim P_change As Variant
    Dim tot_volume As Variant
    Dim k As Long
    Dim Y_open As Variant
    Dim Y_close As Variant
      
    'MODERATE PART
    For Each ws In Worksheets
        'set column headers
        ws.Range("I1") = "Ticker"
        ws.Range("j1") = "Yearly Change"
        ws.Range("k1") = "Percent Change"
        ws.Range("l1") = "Total Stock Volume"
        ws.Columns("I:L").AutoFit
        'set initial variables
        lastrowdata = ws.Cells(Rows.Count, 1).End(xlUp).Row
        k = 2
        Y_open = ws.Cells(2, 3)
        
        For i = 2 To lastrowdata
            current_ticker = ws.Cells(i, 1)
            next_ticker = ws.Cells(i + 1, 1)
            'If first Y_open in ticker is 0 this then increments the starting Y_open to the first value greater then 0
            If Y_open = 0 Then
                Y_open = ws.Cells(i, 3)
            End If
            If current_ticker <> next_ticker Then
                'Plot ticker out to cell
                ws.Cells(k, 9) = current_ticker
                'Plot  yearly change out to cell
                Y_close = ws.Cells(i, 6)
                Y_change = Y_close - Y_open
                ws.Cells(k, 10) = Y_change
                'Plot percent change out to cell
                If Y_close = 0 And Y_open = 0 Then
                    P_change = 0
                Else
                    P_change = (Y_close - Y_open) / Y_open
                End If
                ws.Cells(k, 11) = P_change
                ws.Cells(k, 11).NumberFormat = "0.00%"
                If P_change < 0 Then
                    ws.Cells(k, 11).Interior.ColorIndex = 3
                ElseIf P_change >= 0 Then
                    ws.Cells(k, 11).Interior.ColorIndex = 4
                End If
                ' Store next Y_open value
                Y_open = ws.Cells(i + 1, 3)
                'Plot total stock volume out to cell
                tot_volume = tot_volume + ws.Cells(i, 7)
                ws.Cells(k, 12) = tot_volume
                k = k + 1
                tot_volume = 0
            Else
                tot_volume = tot_volume + ws.Cells(i, 7)
            End If
        Next i
        
        'HARD PART
        Dim Greatest_Increase As Double
        Dim Greatest_Decrease As Double
        Dim Greatest_Tot_Vol As Variant
        Dim Greatest_Increase_Ticker As String
        Dim Greatest_Decrease_Ticker As String
        Dim Greatest_Tot_Vol_Ticker As String
        
        lastrowsummary = ws.Cells(Rows.Count, 9).End(xlUp).Row
        
        For j = 2 To lastrowsummary
            'Check for greatest increase
            If ws.Cells(j, 11) > 0 Then
                If ws.Cells(j, 11) > Greatest_Increase Then
                    Greatest_Increase = ws.Cells(j, 11)
                    Greatest_Increase_Ticker = ws.Cells(j, 9)
                End If
            End If
            'check for greatest decrease
            If ws.Cells(j, 11) < 0 Then
                If ws.Cells(j, 11) < Greatest_Decrease Then
                    Greatest_Decrease = ws.Cells(j, 11)
                    Greatest_Decrease_Ticker = ws.Cells(j, 9)
                End If
            End If
            'check for greatest total volume
            If ws.Cells(j, 12) > Greatest_Tot_Vol Then
                Greatest_Tot_Vol = ws.Cells(j, 12)
                Greatest_Tot_Vol_Ticker = ws.Cells(j, 9)
            End If
        Next j
        
        'Set column and row headers
        ws.Cells(2, 14) = "Greatest % Increase"
        ws.Cells(3, 14) = "Greatest % Decrease"
        ws.Cells(4, 14) = "Greatest Total Volume"
        ws.Cells(1, 15) = "Ticker"
        ws.Cells(1, 16) = "Value"
        'Output values to spreadsheet
        ws.Cells(2, 15) = Greatest_Increase_Ticker
        ws.Cells(2, 16) = Greatest_Increase
        ws.Cells(3, 15) = Greatest_Decrease_Ticker
        ws.Cells(3, 16) = Greatest_Decrease
        ws.Cells(4, 15) = Greatest_Tot_Vol_Ticker
        ws.Cells(4, 16) = Greatest_Tot_Vol
        ws.Columns("n:p").AutoFit
        ws.Range("p2:p3").NumberFormat = "0.00%"
        'reset variables
        Greatest_Increase_Ticker = ""
        Greatest_Increase = 0
        Greatest_Decrease_Ticker = ""
        Greatest_Decrease = 0
        Greatest_Tot_Vol_Ticker = ""
        Greatest_Tot_Vol = 0
    Next ws
End Sub

