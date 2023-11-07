Sub WorksheetLoop_YearlyChange()
         Dim worksheetCount As Integer
         Dim I As Integer
         Dim j As Long
         Dim numTicker As Integer
         Dim yearlyChange As Double
         Dim percentage As Double
         Dim openingPrice As Double
         Dim closingPrice As Double
         Dim stockVolume As Double
         
         worksheetCount = ActiveWorkbook.Worksheets.Count
         
         For I = 1 To worksheetCount
                ActiveWorkbook.Worksheets(I).Activate
                ticker_row = 2
                yearlyChange = 0
                percentage = 0
                openingPrice = 0
                numTicker = 0
                stockVolume = 0
                Cells(1, "J").Value = "ticker_list"
                Cells(1, "K").Value = "yearlyChange_test"
                Cells(1, "L").Value = "percentage"
                Cells(1, "M").Value = "stockVolume"
                
                For j = 2 To ActiveWorkbook.Worksheets(I).UsedRange.Rows.Count
                        If openingPrice = 0 Then
                        openingPrice = Cells(j, 3).Value
                        End If
                        
                        stockVolume = stockVolume + Cells(j, 7).Value
                
                        If Cells(j + 1, 1).Value <> Cells(j, 1).Value Then
                        ticker = Cells(j, 1).Value
                        Cells(ticker_row, "J").Value = ticker
                        numTicker = numTicker + 1
                        Cells(numTicker + 1, 10) = Cells(j, 1).Value
                        closingPrice = Cells(j, 6)
                        yearlyChange = closingPrice - openingPrice
                        Cells(numTicker + 1, 11).Value = yearlyChange
                        
            If yearlyChange > 0 Then
                Cells(numTicker + 1, 11).Interior.ColorIndex = 4
            ElseIf yearlyChange < 0 Then
                Cells(numTicker + 1, 11).Interior.ColorIndex = 3
            Else
                Cells(numTicker + 1, 11).Interior.ColorIndex = 6
            End If
                        
                        If openingPrice = 0 Then
                        percentage = 0
                        Else
                        percentage = (yearlyChange / openingPrice)
                        End If
                        Cells(numTicker + 1, 12).Value = Format(percentage, "Percent")
                        
                        Cells(numTicker + 1, 13).Value = stockVolume
                        End If
                                   
            If percentage > 0 Then
           Cells(numTicker + 1, 12).Interior.ColorIndex = 4
            ElseIf percentage < 0 Then
            Cells(numTicker + 1, 12).Interior.ColorIndex = 3
            Else
            Cells(numTicker + 1, 12).Interior.ColorIndex = 6
            End If

                Next j

         Next I
         
      End Sub
