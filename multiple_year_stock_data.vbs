Sub yearstocks()

    'For going through all worksheets
    For Each ws In Worksheets
    
    'Entering in for the outputs
    ws.Range("I1").Value = "Ticker"
    ws.Range("J1").Value = "Yearly Change"
    ws.Range("K1").Value = "Percent Change"
    ws.Range("L1").Value = "Total Stock Value"
    ws.Range("O2").Value = "Greatest % Increase"
    ws.Range("O3").Value = "Greatest % Decrease"
    ws.Range("O4").Value = "Greatest Total Volaume"
    ws.Range("P1").Value = "Ticker"
    ws.Range("Q1").Value = "Value"
    
    'Defining variables
    Dim yearopen As Double
    Dim yearend As Double
    Dim yearchange As Double
    Dim percent As Double
    Dim greatinc As Double
    Dim greatdec As Double
    Dim lastrow As Long
    Dim ticker As String
    Dim ticker_vol As Double
    Dim tickercounter As Long
    Dim rowstart As Long
    Dim rowend As Long
    Dim open_amt As Long
    
    'Defining the ticker
    ticker = " "
    ticker_vol = 0
    tickercounter = 2
    j = 2
    
    'Defining where these variables will start
    greatinc = 0
    greatdec = 0
    rowstart = 2
    open_amt = 2
    
    'Getting the ticker
    lastrow = ws.Cells(Rows.Count, 1).End(xlUp).Row
        
        For i = 2 To lastrow
        
        ticker_vol = ticker_vol + ws.Cells(i, 7).Value
        If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then
    
        ticker = ws.Cells(i, 1).Value
        ws.Range("I" & rowstart).Value = ticker
        ws.Range("L" & rowstart).Value = ticker_vol
        ticker_vol = 0
        
        'Defining these year change variables
        yearend = ws.Range("F" & i)
        yearopen = ws.Range("C" & open_amt)
        yearchange = yearend - yearopen
        ws.Range("J" & rowstart).Value = yearchange
        
        If yearopen = 0 Then
            percent = 0
        Else
            yearopen = ws.Range("C" & open_amt)
            percent = yearchange / yearopen
        End If
        
        ws.Range("K" & rowstart).NumberFormat = "0.00%"
        ws.Range("K" & rowstart).Value = percent
        
        'Conditional formatting
        If ws.Range("J" & rowstart).Value >= 0 Then
            ws.Range("J" & rowstart).Interior.ColorIndex = 4
        Else
            ws.Range("J" & rowstart).Interior.ColorIndex = 3
        End If
        
        rowstart = rowstart + 1
        open_amt = i + 1
        End If
    Next i
    
    
    'This is for "Greatest % Inc/Dec. Getting column & row. Like alphabetical testing, doing Cells broke Excel
        For i = 2 To lastrow
            If ws.Range("K" & i).Value > ws.Range("Q2").Value Then
                ws.Range("Q2").Value = ws.Range("K" & i).Value
                ws.Range("P2").Value = ws.Range("I" & i).Value
            End If
            
        If ws.Range("K" & i).Value < ws.Range("Q3").Value Then
            ws.Range("Q3").Value = ws.Range("K" & i).Value
                    ws.Range("P3").Value = ws.Range("I" & i).Value
            End If

        If ws.Range("L" & i).Value > ws.Range("Q4").Value Then
                    ws.Range("Q4").Value = ws.Range("L" & i).Value
                    ws.Range("P4").Value = ws.Range("I" & i).Value
            End If

            Next i
            '2 decimal places like the screenshots in Canvas
        ws.Range("Q2").NumberFormat = "0.00%"
        ws.Range("Q3").NumberFormat = "0.00%"
     
    Next ws
    
End Sub