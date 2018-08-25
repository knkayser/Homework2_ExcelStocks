Attribute VB_Name = "Module3"
Sub TickerLoop2016()

'Declare variables
Dim Ticker As String
Dim TotalStockVolume As Double
Dim Summary_Table_Row As Integer
Dim YearlyChange As Double
Dim OpeningValue As Double
Dim ClosingValue As Double
Dim PercentChange As Double

TotalStockVolume = 0
OpeningValue = Cells(2, 3).Value

'Set up summary table
Summary_Table_Row = 2

lastRow = Worksheets("2016").Cells(Rows.Count, "A").End(xlUp).Row + 1

For i = 2 To lastRow
        
        If Cells(i + 1, 1).Value <> Cells(i, 1).Value Then
        
            'Assign ticker value to Ticker
            Ticker = Cells(i, 1).Value
            
            'Add to running total
            TotalStockVolume = TotalStockVolume + Cells(i, 7).Value
            
            'Print Ticker in correct column
            Range("I" & Summary_Table_Row).Value = Ticker
            
            'Print running total to table
            Range("L" & Summary_Table_Row).Value = TotalStockVolume
            
            'Reset stock volume
            TotalStockVolume = 0
            
            'Determined ending value
            ClosingValue = Cells(i, 6).Value
        
            'Calculate difference
            YearlyChange = ClosingValue - OpeningValue
        
            'Send value to table
            Range("J" & Summary_Table_Row).Value = YearlyChange
            
                If OpeningValue = 0 Then
                    PercentChange = 0
                
                Else
            
                'Calculate percentage change
                PercentChange = YearlyChange / OpeningValue
    
                'Reset opening value
                OpeningValue = Cells(i + 1, 3).Value
            
            End If
            
            'Send value to table
            Range("K" & Summary_Table_Row).Value = PercentChange
            
            'Add to summary table
            Summary_Table_Row = Summary_Table_Row + 1
            
        Else
            
            'Add to running total
            TotalStockVolume = TotalStockVolume + Cells(i, 7).Value
   
        End If
        
    Next i

'Hard Option
Dim LargestPercent As Double
Dim Smallestercent As Double
Dim LargestVolume As Double
Dim StartingPercent1 As Double
Dim StartingPercent2 As Double
Dim Ticker_New_1 As String
Dim Ticker_New_2 As String
Dim Ticker_New_3 As String

'Create For loop to search for greatest % increase

StartingPercent1 = Cells(2, 11).Value

lastRow2 = Worksheets("2016").Cells(Rows.Count, "J").End(xlUp).Row + 1

For i = 2 To lastRow2

    If StartingPercent1 < Cells(i, 11).Value Then
        LargestPercent = Cells(i, 11).Value
        StartingPercent1 = LargestPercent
        Ticker_New_1 = Cells(i, 9).Value
        
    End If
    
Next i

'Create For loop to search for smallest %

StartingPercent2 = Cells(2, 11).Value
    
For i = 2 To lastRow2
    
    If StartingPercent2 > Cells(i, 11).Value Then
        SmallestPercent = Cells(i, 11).Value
        StartingPercent2 = SmallestPercent
        Ticker_New_2 = Cells(i, 9).Value
        
    End If
    
Next i

'Create For loop to search for largest volume

LargestVolume = 0
    
For i = 2 To lastRow2
    
    If Cells(i, 12).Value > LargestVolume Then
        LargestVolume = Cells(i, 12).Value
        Ticker_New_3 = Cells(i, 9).Value
        
    End If

Next i

        Cells(2, 17).Value = StartingPercent1
        Cells(3, 17).Value = SmallestPercent
        Cells(4, 17).Value = LargestVolume
        Cells(2, 16).Value = Ticker_New_1
        Cells(3, 16).Value = Ticker_New_2
        Cells(4, 16).Value = Ticker_New_3
        

End Sub




