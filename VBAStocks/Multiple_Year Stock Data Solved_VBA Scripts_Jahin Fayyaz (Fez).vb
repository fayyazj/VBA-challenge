'call all sub-routines

Sub YearStockData():

    Ticker
    Yearly_Change
    Percent_Change
    Total_Stock_Volume
    Greatest
    
End Sub



'This sub is to automate the call of tickers and remove any duplications from column A to column I
    
Sub Ticker():
    
    Range("I1").Value = "Ticker"
        
    'Define last row of sheet
    Dim LastRow As Long
    Dim sht As Worksheet
    Set sht = ActiveSheet
    LastRow = 1 + sht.Cells(sht.Rows.Count, "A").End(xlUp).Row
        
    'Define New ticker and Old ticker so that the for loop can compare
    Dim OldTicker As String
    Dim NewTicker As String
        
    'Define next row so that tickers populate in the succeeding row when new ticker does not equal the old ticker
    Dim NewRow As Long
        
    'Initial Variable Assignment
    OldTicker = Cells(2, 1).Value
    NewRow = 2
        
    'For Loop, Loop through first columnand when a new ticker is reached, output will go in next row in column i
                    
        For i = 2 To LastRow
            NewTicker = Cells(i, 1).Value
                
            If NewTicker <> OldTicker Then 'Comparing to see if there is a new ticker versus the old ticker
                Cells(NewRow, 9).Value = OldTicker
                OldTicker = NewTicker
                NewRow = NewRow + 1
                    
            End If
                
        Next i

End Sub
    
'Yearly change from opening price at the beginning of a given year to the closing price at the end of that year.

Sub Yearly_Change():


    Range("J1").Value = "Yearly Change"
    
    'Format decimals in column K to percent/to two decimals
    Range("J:J").NumberFormat = "0.000000000"
    
    'Define last row of sheet
    Dim LastRow As Long
    Dim sht As Worksheet
    Set sht = ActiveSheet
    LastRow = 1 + sht.Cells(sht.Rows.Count, "A").End(xlUp).Row
        
    'Define Yearly Change.
    Dim YearlyChange As Double
        
    'Define New ticker and Old ticker so that the for loop can compare
    Dim OldTicker As String
    Dim NewTicker As String
        
    'Define next row so that tickers populate in the new row when new ticker does not equal the old ticker
    Dim NewRow As Long
            
    'Define opening and closing prices
    Dim OpeningPrice As Double
    Dim ClosingPrice As Double
    
    
    'Initial Variable Assignment
    OldTicker = Cells(2, 1).Value
    NewRow = 2
    OpeningPrice = Range("C2").Value
        
    'For Loop, Loop through first columnand when a new ticker is reached, output will go in next row in column i
                    
        For i = 2 To LastRow
            NewTicker = Cells(i, 1).Value
                
            If NewTicker <> OldTicker Then 'Comparing to see if there is a new ticker versus the old ticker
                ClosingPrice = Cells(i - 1, 6)
                YearlyChange = ClosingPrice - OpeningPrice
                Cells(NewRow, 10).Value = YearlyChange
                
                'Conditional Color Formatting
                If YearlyChange > 0 Then
                    Cells(NewRow, 10).Interior.ColorIndex = 4
                ElseIf YearlyChange < 0 Then
                    Cells(NewRow, 10).Interior.ColorIndex = 3
                End If
                
                OldTicker = NewTicker
                NewRow = NewRow + 1
                OpeningPrice = Cells(i, 3).Value
            End If
                
        Next i
    
End Sub

Sub Percent_Change():

Range("K1").Value = "Percent Change"

'Format decimals in column K to percent/to two decimals
Range("K:K").NumberFormat = "0.00%"
    
    'Define last row of sheet
    Dim LastRow As Long
    Dim sht As Worksheet
    Set sht = ActiveSheet
    LastRow = 1 + sht.Cells(sht.Rows.Count, "A").End(xlUp).Row
        
    'Define Yearly Change
    Dim YearlyChange As Double
    
    'Define Percent Change
    Dim PercentChange As Double
        
    'Define New ticker and Old ticker so that the for loop can compare
    Dim OldTicker As String
    Dim NewTicker As String
        
    'Define next row so that tickers populate in the succeeding row when new ticker does not equal the old ticker
    Dim NewRow As Long
            
    'Define opening and closing prices
    Dim OpeningPrice As Double
    Dim ClosingPrice As Double
    
    
    'Initial Variable Assignment
    OldTicker = Cells(2, 1).Value
    NewRow = 2
    OpeningPrice = Range("C2").Value
        
    'For Loop, Loop through first columnand when a new ticker is reached, output will go in next row in column i
                    
        For i = 2 To LastRow
            NewTicker = Cells(i, 1).Value
                
            If NewTicker <> OldTicker Then 'Comparing to see if there is a new ticker versus the old ticker
                ClosingPrice = Cells(i - 1, 6)
                YearlyChange = ClosingPrice - OpeningPrice
                ' to prevent errors when dividing by 0, if there is a 0 in the dataset
                If OpeningPrice = 0 Then
                    PercentChange = 0
                Else
                    PercentChange = YearlyChange / OpeningPrice
                End If
                
                Cells(NewRow, 11).Value = PercentChange
                OldTicker = NewTicker
                NewRow = NewRow + 1
                OpeningPrice = Cells(i, 3).Value
                
            End If
                
        Next i

End Sub
    
Sub Total_Stock_Volume()

   'The total stock volume of the stock. Outputs to column L.
   Range("L1").Value = "Total Stock Volume"
   
   'Variable Declarations
   Dim LastRow, NewRow, TotalVolume, NextVolume As Long
   Dim OldTicker, NewTicker As String
   Dim sht As Worksheet
   
   'Setting LastRow
   Set sht = ActiveSheet
   LastRow = 1 + sht.Cells(sht.Rows.Count, "A").End(xlUp).Row
   
   'Variable Initialization
   OldTicker = Cells(2, 1).Value
   NewRow = 2
   TotalVolume = 0
   
   'program loop
   For i = 2 To LastRow
       NewTicker = Cells(i, 1).Value
       
       If NewTicker <> OldTicker Then 'Compares the values of NewTicker and OldTicker
           Cells(NewRow, 12).Value = TotalVolume
           OldTicker = NewTicker
           NewRow = NewRow + 1
           TotalVolume = Cells(i, 7).Value
       End If
       
       Volume = Cells(i, 7).Value
       TotalVolume = TotalVolume + Volume
       
   Next i
   
End Sub

Sub Greatest():
    'Set LastRow
     Set sht = ActiveSheet
     LastRow = sht.Cells(sht.Rows.Count, "I").End(xlUp).Row
     
     'Define variables
    Dim GLR, GPI, GPD, GTV, NPC, NTV As Double
    
    Range("P1").Value = "Ticker"
    Range("Q1").Value = "Value"
    Range("O2").Value = "Greatest % Increase"
    Range("O3").Value = "Greatest % Decrease"
    Range("O4").Value = "Greatest Total Volume"
    
    'Format decimals in column K to percent/to two decimals
    Range("Q2:Q3").NumberFormat = "0.00%"
    Range("Q4").NumberFormat = "00000000000"
    
    GPI = 0 'Greatest Percent Increase
    GPD = 0 'Greatest Percent Decrease
    GTV = 0 'Greatest Total Volume
    
    For i = 2 To LastRow
    
        'New Percent Change
        NPC = Range("K" & i).Value
        
        'New Total Volume
        NTV = Range("L" & i).Value
        
        'Compare all the percent volume and find greatest increase
        If NPC > GPI Then
            GPI = NPC
            Range("P2").Value = Range("I" & i).Value
            Range("Q2").Value = GPI
            
        'Compare all percent volume and find greatest decrease
        ElseIf NPC < GPD Then
            GPD = NPC
            Range("P3").Value = Range("I" & i).Value
            Range("Q3").Value = GPD 'can also be Range("Q", i).Value
        End If
        
        'compare greatest total volume
        If NTV > GTV Then
            GTV = NTV
            Range("P4").Value = Range("I" & i).Value
            Range("Q4").Value = GTV
            
        End If
        
    Next i
    
End Sub

'Loops through all the worksheets. kept in its own sub to prevent potential crashes

Sub LoopThrough():

    For Each WS In ActiveWorkbook.Worksheets
        WS.Activate
        YearStockData
    Next WS
    
End Sub
