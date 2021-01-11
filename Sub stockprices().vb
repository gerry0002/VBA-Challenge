Sub stockprices()

    Dim lastRow As Double
    Dim ticker As String
    Dim Ychange As Double
    Dim PChart As Double
    Dim TStcVol As Double
    Dim begYear As Double
    Dim endYear As Double
    Dim CurrStock As Double
    Dim GrtInc As Double
    Dim GrtDec As Double
    Dim GrtTot As Double
    Dim GIT As String
    Dim GDT As String
    Dim GTV As String
    
    lastRow = Cells(Rows.Count, 1).End(xlUp).Row + 1
    
    Range("I1").Value = "Ticker"
    Range("J1").Value = "Yearly Change"
    Range("K1").Value = "Percent Chart"
    Range("L1").Value = "Total Stock Volume"
    
    Range("O2").Value = "Greatest % Increase"
    Range("O3").Value = "Greatest % Decrease"
    Range("O4").Value = "Greatest Total volume"
    
    Range("P1").Value = "Ticker"
    Range("Q1").Value = "Value"
    
    
    CurrStock = 0
       
    ticker = ""
    
    For i = 2 To lastRow
    
        
        If ticker = Cells(i, 1).Value Then
        
            TStcVol = TStcVol + Cells(i, 7)
            
        Else
        
            CurrStock = CurrStock + 1
            If i <> 2 Then
                Cells(CurrStock, 9).Value = ticker
            End If
            
            If ticker <> "" Then
                endYear = Cells((i - 1), 6).Value
                Cells(CurrStock, 10).Value = endYear - begYear
                If (endYear - begYear) < 0 Then
                    Cells(CurrStock, 10).Interior.ColorIndex = 3
                Else
                    Cells(CurrStock, 10).Interior.ColorIndex = 4
                End If
                
                
                Cells(CurrStock, 12) = TStcVol
                
                If begYear <> 0 Then
                    Cells(CurrStock, 11) = (endYear - begYear) / begYear
                
                    If GrtInc < (endYear - begYear) / begYear Then
                        GrtInc = (endYear - begYear) / begYear
                        GIT = ticker
                    End If
                    
                    If GrtDec > (endYear - begYear) / begYear Then
                        GrtDec = (endYear - begYear) / begYear
                        GDT = ticker
                    End If
                    
                    If GrtTot < TStcVol Then
                        GrtTot = TStcVol
                        GTV = ticker
                    End If
                
                Else
                    Cells(CurrStock, 11) = 0
                End If
                
                
            End If
            begYear = Cells(i, 3).Value
            ticker = Cells(i, 1).Value
            TStcVol = 0
            
            
        End If
     
        
    
    Next i
    
    Range("P2") = GIT
    Range("P3") = GDT
    Range("P4") = GTV
    
    Range("Q2") = GrtInc
    Range("Q3") = GrtDec
    Range("Q4") = GrtTot
    
    
    
    Range("K:K").NumberFormat = "0.00%"
    Range("Q2:Q3").NumberFormat = "0.00%"
    

End Sub