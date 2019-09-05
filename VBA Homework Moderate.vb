Sub VBA_hw_moderate()

    For Each ws In Worksheets
    ws.Activate

        Range("I1").Value = "Ticker"
        Range("J1").Value = "Yearly Change"
        Range("K1").Value = "Percent Change"
        Range("L1").Value = "Total Stock Volume"
        
        Range("J:K").EntireColumn.AutoFit
        Range("L:L").EntireColumn.AutoFit

        Dim Ticker_Name As String


        Dim Open_Date As Boolean
        Open_Date = True

        Dim Open_Price As Double
        Dim End_Price As Double
        
        Dim Yearly_Change As Double
        Dim Percent_Change As Double

        Dim Total_Volume As Double
        Total_Volume = 0

        Dim Ticker_Table_Row As Integer
        Ticker_Table_Row = 2

        Dim End_Row As Long

        End_Row = ws.Cells(Rows.Count, 1).End(xlUp).Row

        For i = 2 To End_Row

            If Cells(i + 1, 1).Value <> Cells(i, 1).Value Then
                
                Ticker_Name = Cells(i, 1).Value
                
                Total_Volume = Total_Volume + Cells(i, 7).Value
                
                Range("I" & Ticker_Table_Row).Value = Ticker_Name

                Range("L" & Ticker_Table_Row).Value = Total_Volume
                
                End_Price = Cells(i, 6).Value
                Yearly_Change = End_Price - Open_Price
                Range("J" & Ticker_Table_Row).Value = Yearly_Change
                
                        
                Percent_Change = (Yearly_Change / Open_Price)
                Range("K" & Ticker_Table_Row).Value = Percent_Change
                Range("K" & Ticker_Table_Row).NumberFormat = "0.00%"


                If Range("J" & Ticker_Table_Row).Value > 0 Then
                    Range("J" & Ticker_Table_Row).Interior.ColorIndex = 4
                Else
                    Range("J" & Ticker_Table_Row).Interior.ColorIndex = 3
                End If
                
                                
                Ticker_Table_Row = Ticker_Table_Row + 1
                
                Open_Date = True

                Total_Volume = 0
                
            Else
                
                Total_Volume = Total_Volume + Cells(i, 7).Value
                
                If Open_Date And Cells(i, 3).Value <> 0 Then
                    Open_Price = Cells(i, 3).Value
                    Open_Date = False
                
                End If
                
                    
            End If
            
        Next i
        
    Next ws

End Sub