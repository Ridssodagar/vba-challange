Attribute VB_Name = "Module1"

Sub alphabetical_testing()

    Dim ws As Worksheet
    
    
    'Creating Variables and assigning Values
    
    Dim Ticker As String
    Dim openingrow As Long
    
    
    Dim yearlychange As Double
    Dim percentchange As Double
    
    Dim RowCount As Long
    
    
    Dim Tickertotal As Double
    Dim outputrow As Long
    Dim inputrow As Long

    Dim greatestincrease As Double
    Dim greatestdecrease As Double
    Dim greatesttotalvolume As Double
    
'loops through all the stocks worksheets
    
    For Each ws In Worksheets
        openingrow = 2
        outputrow = 2
        Tickertotal = 0
        yearlychange = 0
        greatestincrease = 0
        greatestdecrease = 0
        greatesttotalvolume = 0
        
        'assigning colums headers
        
        ws.Range("I1").Value = "Ticker"
        ws.Range("J1").Value = "Yearly Change"
        ws.Range("K1").Value = "Percent Change"
        ws.Range("L1").Value = "Total Stock Volume"
        ws.Range("P1").Value = "Ticker"
        ws.Range("Q1").Value = "Value"
        ws.Range("O2").Value = "Greatest%Increase"
        ws.Range("O3").Value = "Greatest%decrease"
        ws.Range("O4").Value = "Greatest Total Volume"
        
        

        'Row Count
        
        RowCount = ws.Range("A1").End(xlDown).Row
        
        'looping all ticker
        
        For inputrow = 2 To RowCount
            Ticker = ws.Cells(inputrow, 1).Value
            If ws.Cells(inputrow + 1, 1).Value <> Ticker Then
            
                Tickertotal = Tickertotal + ws.Cells(inputrow, 7).Value
                
                
                
''
''            If ws.Cells(openingvalue, 3) = 0 Then
''
''                For findvalue = openingvalue To i
''
''                    If ws.Cells(findvalue, 3).Value <> 0 Then
''
''                        openingrow = findvalue
''
''                        Exit For
''                    End If
''
''                Next findvalue
''
''            End If

                 'calculations
                
                 yearlychange = ws.Cells(inputrow, 6) - ws.Cells(openingrow, 3)
                 percentchange = (yearlychange / ws.Cells(openingrow, 3))
            
           
                 'output
            
                 ws.Range("I" & outputrow).Value = Ticker
                
                 ws.Range("L" & outputrow).Value = Tickertotal
            
                 ws.Range("J" & outputrow).Value = yearlychange
                 
                  If ws.Range("J" & outputrow).Value > 0 Then
                  ws.Range("J" & outputrow).Interior.ColorIndex = 4

                  ElseIf ws.Range("J" & outputrow).Value < 0 Then
                  ws.Range("J" & outputrow).Interior.ColorIndex = 3

            End If
        
                 ws.Range("K" & outputrow).Value = FormatPercent(percentchange, 2)
            
                'prepare for next row
                 openingrow = inputrow + 1
                 Tickertotal = 0
                 yearlychange = 0
                 outputrow = outputrow + 1
        
            Else
                 Tickertotal = Tickertotal + ws.Cells(inputrow, 7).Value
                
            End If
            
            
        Next inputrow
        
        'output greatest & least value
           
        For i = 2 To RowCount
        
            If ws.Cells(i, 11).Value > greatestincrease Then
                greatestincrease = ws.Cells(i, 11).Value
                
                ws.Range("Q2").Value = FormatPercent(greatestincrease)
                ws.Range("P2").Value = Ticker
                ws.Range("P2").Value = ws.Cells(i, 9).Value
                
         
           End If
           
           Next i
           
        For j = 2 To RowCount
        
            If ws.Cells(j, 11).Value < greatestdecrease Then
                greatestdecrease = ws.Cells(j, 11).Value
                
                ws.Range("Q3").Value = FormatPercent(greatestdecrease)
                ws.Range("P3").Value = Ticker
                ws.Range("P3").Value = ws.Cells(j, 9).Value
                
            End If
            
            Next j
            
        For k = 2 To RowCount
        
            If ws.Cells(k, 12).Value > greatesttotalvolume Then
                greatesttotalvolume = ws.Cells(k, 12).Value
                
                ws.Range("Q4").Value = greatesttotalvolume
                ws.Range("P4").Value = Ticker
                ws.Range("P4").Value = ws.Cells(k, 9).Value
                
            End If
            
            Next k
            
            

Next ws



End Sub



