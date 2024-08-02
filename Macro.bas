Attribute VB_Name = "Module1"
Sub ticker()

'initiate worksheet reading
For Each ws In Worksheets

'variable fortickers to be identified
Dim ticker As String

Dim volume As Double
        volume = 0

'variable to define row ticker will occupy
Dim row As Integer
        row = 2

'label headers for analysis to go in
ws.Cells(1, 9).Value = "Ticker"
ws.Cells(1, 10).Value = "Quarterly Change"
ws.Cells(1, 11).Value = "Percent Change"
ws.Cells(1, 12).Value = "Total Stock Volume"

'variables for percent and quarterly change, and close price of ticker
Dim percent_change, quarterly_change, closeprice As Double

'variable of open price of ticker
Dim openprice As Double
'Define variables value
openprice = ws.Cells(2, 3).Value

'define lastrow
Dim LastRow As Long

'find last row with ticker data
LastRow = ws.Cells(Rows.Count, 1).End(xlUp).row

'loop through each row
For i = 2 To LastRow

'volume variable to calculate total tickers held

        'if ticker is different than one before
        If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then
        
            'get ticker of previous before name changes
            ticker = ws.Cells(i, 1).Value
            
            'place name in ticker column
            ws.Range("I" & row).Value = ticker
            
            'add volume of stock trade that has been added up till now with last volume count
            volume = volume + ws.Cells(i, 7).Value
            
            'place volume of stock into volume column
            ws.Range("L" & row).Value = volume
            
            'Define last close price of year
            closeprice = ws.Cells(i, 6).Value
            
            'use closeprice to get quarterly change from open price of start date
            quarterly_change = (closeprice - openprice)
            
            'place value in quarterly column
            ws.Range("J" & row).Value = quarterly_change
            
            'calculate percent change
            percent_change = (((quarterly_change / openprice) * 100) & "%")
            
            'place value in percent change column
            ws.Range("K" & row).Value = percent_change
            
                 'Conditional formating
                    If ws.Cells(i, 10).Value >= 0 Then
                
                    'Set cell background color to green
                    ws.Cells(i, 10).Interior.ColorIndex = 4
                
                    Else
                
                    'Set cell background color to red
                    ws.Cells(i, 10).Interior.ColorIndex = 3
                
                    End If
            
            'change rowto move down list
            row = row + 1
            
            'volume reset to zero
            volume = 0
            
            'change the opening price based on row i that changes to new ticker
            openprice = ws.Cells(i + 1, 3)
            
        Else
            'when tickers are the same add the volume of stock to get total
            volume = volume + ws.Cells(i, 7).Value
            
            If ws.Cells(i, 10).Value >= 0 Then
                
                    'Set cell background color to green
                    ws.Cells(i, 10).Interior.ColorIndex = 4
                
                    Else
                
                    'Set cell background color to red
                    ws.Cells(i, 10).Interior.ColorIndex = 3
                
                    End If
            
        End If
Next i

'label columns for summary of data
ws.Cells(2, 15).Value = "Greatest % Increase"
ws.Cells(3, 15).Value = "Greatest % Decrease"
ws.Cells(4, 15).Value = "Greatest Total Volume"
ws.Cells(1, 16).Value = "Ticker"
ws.Cells(1, 17).Value = "Value"

'last row determination
LastRow = ws.Cells(Rows.Count, 9).End(xlUp).row

'determine min max value
GreatestVolume = ws.Cells(2, 12).Value
GreatestPercent = ws.Cells(2, 11).Value
NegativePercent = ws.Cells(2, 11).Value
'find max of total volume
For i = 2 To LastRow
        
            'Find the maximum percent change
            If ws.Cells(i, 11).Value > GreatestPercent Then
                GreatestPercent = ws.Cells(i, 11).Value
                ws.Cells(2, 16).Value = ws.Cells(i, 9).Value
                
                Else
                
                GreatestPercent = GreatestPercent
                
                End If
            'Find the minimum percent change
            If ws.Cells(i, 11).Value < NegativePercent Then
                NegativePercent = ws.Cells(i, 11).Value
                ws.Cells(3, 16).Value = ws.Cells(i, 9).Value
                
                Else
                
                NegativePercent = NegativePercent
                
                End If
            'Find the maximum volume of trade
            If ws.Cells(i, 12).Value > GreatestVolume Then
                GreatestVolume = ws.Cells(i, 12).Value
                ws.Cells(4, 16).Value = ws.Cells(i, 9).Value
            Else
                
                GreatestVolume = GreatestVolume
                
                End If
            ws.Cells(2, 17).Value = Format(GreatestPercent, "Percent")
            ws.Cells(3, 17).Value = Format(NegativePercent, "Percent")
            ws.Cells(4, 17).Value = Format(GreatestVolume, "Scientific")
        Next i
Next ws

End Sub
