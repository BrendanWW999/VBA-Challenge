Attribute VB_Name = "Module1"
Sub MultiSheets()

Dim xSh As Worksheet
    For Each xSh In ThisWorkbook.Worksheets
        xSh.Select
        Call Stockmarket
    Next xSh

End Sub

Sub Stockmarket()

    Dim Tickername As String
    Dim Tickerrow As Integer
    Dim Stockvol As Double
    Dim Yearend As Double
    Dim Yearstart As Double
    Dim Yearperc As Double
    Dim RowTotal As Double
    Dim Rowcount As Double
    Dim RowResults As Double
    
    Dim GreIncTick As Double
    Dim GreIncValue As Double
    Dim GreDecTick As Double
    Dim GreDecValue As Double
    Dim GreVolTick As Double
    Dim GreVolValue As Double
    
    Yearperc = 0
    Yearchange = 0
    Stockvol = 0
    Tickerrow = 2
    
    RowTotal = Cells.Rows.Count
    Rowcount = 0
    
    For k = 1 To RowTotal
        If Cells(k, 1).Value = "" Then
            If Rowcount = 0 Then
                Rowcount = k - 1
                k = RowTotal - 1
            End If
        End If
    Next k
        
        Cells(1, 16) = "Ticker"
        Cells(1, 17) = "Value"
        Cells(2, 15) = "Greatest % Increase"
        Cells(3, 15) = "Greatest % Decrease"
        Cells(4, 15) = "Greatest Total Volume"
        Cells(1, 9) = "Ticker"
        Cells(1, 10) = "Yearly Change"
        Cells(1, 11) = "Yearly Change"
        Cells(1, 12) = "Total Stock"
        Columns("J").ColumnWidth = 13
        Columns("K").ColumnWidth = 13
        Columns("L").ColumnWidth = 14
        Columns("O").ColumnWidth = 19
        Columns("Q").ColumnWidth = 14
        
        For i = 2 To Rowcount
            
            If Cells(i, 1) <> Cells(i - 1, 1) Then
            Yearstart = Cells(i, 3).Value
                                   
            End If
                                               
            If Cells(i + 1, 1).Value <> Cells(i, 1).Value Then
            Tickername = Cells(i, 1).Value
            
                Yearend = Cells(i, 6).Value
                
                Stockvol = Stockvol + Cells(i, 7).Value
                
                Range("I" & Tickerrow).Value = Tickername
                       
                Range("J" & Tickerrow).Value = Yearend - Yearstart
    
                If Yearstart = 0 Then
                
                    Range("K" & Tickerrow).Value = "0"
                
                    Yearperc = 0
                
                ElseIf Yearend = 0 Then
                
                    Range("K" & Tickeroow).Value = "0"
                
                    Yearperc = 0
               
                Else
                
                    Yearperc = ((1 - (Yearend / Yearstart)) * -1)
                
                End If
                
                    Range("K" & Tickerrow).Value = FormatPercent(Yearperc)
                                   
                    Range("L" & Tickerrow).Value = Stockvol
                
                    Tickerrow = Tickerrow + 1
                
                    Stockvol = 0
            
            Else
            
                Stockvol = Stockvol + Cells(i, 7).Value
           
            End If
        Next i
        
        GreIncTick = 2
        GreIncValue = Cells(2, 11).Value
        GreDecTick = 2
        GreDecValue = Cells(2, 11).Value
        GreVolTick = 2
        GreVolValue = Cells(2, 12).Value
       
        RowResults = 0
        For k = 1 To RowTotal
            If Cells(k, 9).Value = "" Then
                If RowResults = 0 Then
                    RowResults = k - 1
                    k = RowTotal - 1
                End If
            End If
        Next k
        
        For j = 2 To RowResults
            If Cells(j, 10) > 0 Then
                Cells(j, 10).Interior.Color = vbGreen
            ElseIf Cells(j, 10) < 0 Then
                Cells(j, 10).Interior.Color = vbRed
            End If
                      
            If Cells(j, 11).Value > GreIncValue Then
                GreIncValue = Cells(j, 11).Value
                GreIncTick = j
            ElseIf Cells(j, 11).Value < GreDecValue Then
                GreDecValue = Cells(j, 11).Value
                GreDecTick = j
            End If
            
            If Cells(j, 12).Value > GreVolValue Then
                GreVolValue = Cells(j, 12).Value
                GreVolTick = j
            End If
        
        Next j
        
        Cells(2, 16).Value = Cells(GreIncTick, 9).Value
        Cells(3, 16).Value = Cells(GreDecTick, 9).Value
        Cells(4, 16).Value = Cells(GreVolTick, 9).Value
        Cells(2, 17).Value = FormatPercent(GreIncValue)
        Cells(3, 17).Value = FormatPercent(GreDecValue)
        Cells(4, 17).Value = GreVolValue
               

End Sub

