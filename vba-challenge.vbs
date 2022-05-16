Option Explicit
Sub StockLoop()
    Dim TotalVol As Double
    Dim vRow As Long
    Dim i As Long
    Dim position As Long
    Dim OpenPrice As Double
    Dim ClosePrice As Double
    Dim vCount As Integer
    Dim MaxIncrease As Double
    Dim MinIncrease As Double
    Dim MaxVol As Double
    
    Dim Current As Worksheet
    Dim vPercent As Double


    For Each Current In Worksheets
            Current.Activate
            vPercent = 0
            MaxVol = 0
            MaxIncrease = 0
            MinIncrease = 0
            OpenPrice = 0
            ClosePrice = 0
            vCount = 0  'counter to track first row of new ticker
            vRow = ActiveSheet.UsedRange.Rows.Count
            TotalVol = 0
            position = 2     ' row position to populate caculated data
            Range("I1") = "Ticker"
            Range("J1") = "Yearly Changed"
            Range("K1") = "Percent Changed"
            Range("L1") = "Total Stock Volume"
            Range("O2") = "Greatest % Increase"
            Range("O3") = "Greatest % Decrease"
            Range("O4") = "Greatest Total Volume"
            Range("P1") = "Ticker"
            Range("Q1") = "Value"
            For i = 2 To vRow
                If Cells(i + 1, 1) <> Cells(i, 1) Then      'if last row of a current ticket
                    TotalVol = TotalVol + Cells(i, 7)
                    Cells(position, 9) = Cells(i, 1)        'populate ticker name
                    Cells(position, 12) = TotalVol      'populate total Volume
                    ClosePrice = Cells(i, 6)         'update close price from last row
                    
                    If (ClosePrice - OpenPrice) < 0 Then        ' update color
                        Cells(position, 10).Interior.ColorIndex = 3
                    ElseIf (ClosePrice - OpenPrice) = 0 Then
                        Cells(position, 10).Interior.ColorIndex = 6
                    Else
                        Cells(position, 10).Interior.ColorIndex = 4
                    End If
                    Cells(position, 10) = (ClosePrice - OpenPrice)       'populate price change
                    
                    If OpenPrice <> 0 Then  'check if Openprice is not 0
                        vPercent = (ClosePrice - OpenPrice) / OpenPrice
                        Cells(position, 11) = vPercent      'populate percentage change
                    Else
                        'hanlde open price =0 here, don't populate/update percentage change
                        vPercent = 0
                        Cells(position, 11) = 0 ' no data
                    End If
                    
                        
                    Cells(position, 11).NumberFormat = "0.00%"
                    If vPercent > MaxIncrease Then        'update max%increase
                        MaxIncrease = vPercent
                        Cells(2, 16) = Cells(i, 1)
                        Cells(2, 17) = MaxIncrease
                        Cells(2, 17).NumberFormat = "0.00%"
                    ElseIf vPercent < MinIncrease Then      'update min%increase
                        MinIncrease = vPercent
                        Cells(3, 16) = Cells(i, 1)
                        Cells(3, 17) = MinIncrease
                        Cells(3, 17).NumberFormat = "0.00%"
                    End If

                             
                    If (TotalVol > MaxVol) Then     'update Max Volume
                        MaxVol = TotalVol
                        Cells(4, 16) = Cells(i, 1)
                        Cells(4, 17) = MaxVol
                    End If
                                           
                    position = position + 1      'next row position
                    TotalVol = 0        'reset values
                    OpenPrice = 0
                    ClosePrice = 0
                    vCount = 0
                    vPercent = 0
                Else        ' if 1st or middle row of current ticker
                    TotalVol = TotalVol + Cells(i, 7)
                    If vCount = 0 And Cells(i, 3) <> 0 Then       'find the first row with open price <> 0
                        OpenPrice = Cells(i, 3)      'get openprice from first row
                        vCount = vCount + 1         ' indicate first row has been taken data
                    End If
                End If
            Next i
    Next
    
    
End Sub

