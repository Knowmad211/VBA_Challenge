Sub Summarize()

ThisWorkbook.Save


    'define variables
    Dim Ticker As String
    Dim YearlyOpen As Double
    Dim YearlyClose As Double
    Dim Volume As Double
    
    Dim yearlychange As Double
    Dim percentchange As Double
    Dim totalvolume As Double
    Dim summaryrowindex As Integer
    Dim openindex As Long
    Dim greatestincrease As Double
    Dim greatestincreaseticker As String
    Dim greatestdecrease As Double
    Dim greatestdecreaseticker As String
    Dim greatesttotalvolume As Double
    Dim greatesttotalticker As String
    
    Dim i As Long
    
    For Each ws In Worksheets
    
    ws.Range("I1").Value = "Ticker"
    ws.Range("J1").Value = "Yearly Change"
    ws.Range("K1").Value = "Percent Change"
    ws.Range("L1").Value = "Total Stock Volume"
    ws.Range("P1").Value = "Ticker"
    ws.Range("Q1").Value = "Value"
    ws.Range("O2").Value = "Greatest % Increase"
    ws.Range("O3").Value = "Greatest % Decrease"
    ws.Range("O4").Value = "Greatest Total Volume"
    
    summaryrowindex = 2
    openindex = 2
    totalvolume = 0
    greatestincrease = 0
    greatestdecrease = 0
    greatesttotalvolume = 0
    
    LastRow = ws.Cells(Rows.Count, 1).End(xlUp).Row
    
    For i = 2 To LastRow
    
        'add all stock volume together
        totalvolume = totalvolume + ws.Cells(i, 7)
    
    
        'loop through tickers of a given value
        If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then
            
            
            'take the opening value of the first date
            YearlyOpen = ws.Cells(openindex, 3).Value
            
            
            'take the closing value of the last date
            YearlyClose = ws.Cells(i, 6).Value
            
            
            'find the yearly change between the first opening and last closing (closing - opening)
            yearlychange = YearlyClose - YearlyOpen
            
            
            'find percentage change between first opening and last closing ((closing-opening)/opening)
            percentchange = (yearlychange / YearlyOpen)
            
            ws.Range("i" & summaryrowindex).Value = ws.Cells(i, 1).Value
            ws.Range("j" & summaryrowindex).Value = yearlychange
            ws.Range("k" & summaryrowindex).Value = percentchange
            ws.Range("k" & summaryrowindex).NumberFormat = "0.00%"
            ws.Range("l" & summaryrowindex).Value = totalvolume
            
            
            'find the greatest % increase
            If percentchange > greatestincrease Then
                greatestincrease = percentchange
                greatestincreaseticker = ws.Cells(i, 1).Value
            End If
                
            
            'find greates % decrease
            If percentchange < greatestdecrease Then
                greatestdecrease = percentchange
                greatestdecreaseticker = ws.Cells(i, 1).Value
            End If
            
            'find greatest total volume
            If totalvolume > greatesttotalvolume Then
                greatesttotalvolume = totalvolume
                greatesttotalticker = ws.Cells(i, 1).Value
            End If
            
        
            summaryrowindex = summaryrowindex + 1
            openindex = i + 1
            totalvolume = 0
            
        
            
        End If
        
    Next i
            
    ws.Range("P2").Value = greatestincreaseticker
    ws.Range("Q2").Value = greatestincrease
    ws.Range("Q2").NumberFormat = "0.00%"
            
    ws.Range("P3").Value = greatestdecreaseticker
    ws.Range("Q3").Value = greatestdecrease
    ws.Range("Q3").NumberFormat = "0.00%"
            
    ws.Range("P4").Value = greatesttotalticker
    ws.Range("Q4").Value = greatesttotalvolume
        
            
            'across all worksheets
    
  Next ws
    
End Sub
