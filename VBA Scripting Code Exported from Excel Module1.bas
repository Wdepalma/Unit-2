Attribute VB_Name = "Module1"
Sub Stock_Anlysis()

Dim ticker, max_increase_ticker, max_decrease_ticker, max_vol_ticker As String
Dim tcount, scount As Integer
Dim op, cl, maxincrease, maxdecrease As Double
Dim vol, maxvol As LongLong
Dim rcount As Long



scount = ThisWorkbook.Sheets.Count  'COUNT NUMBER OF SHEETS IN WORKBOOK

For j = 1 To scount

    Sheets(j).Activate
    
    'FIND LAST ROW OF ACTIVE SHEET
    rcount = ActiveSheet.Cells(ActiveSheet.Rows.Count, "A").End(xlUp).Row
    
    'WRITE COLUMN HEADERS INTO WORKSHEET
    Cells(1, 9).Value = "Ticker"
    Cells(1, 10).Value = "Yearly Change"
    Cells(1, 11).Value = "Percent Change"
    Cells(1, 12).Value = "Total Stock Volume"
    
    'INITIALIZE VARIABLES FOR EACH SHEET
    tcount = 0                          'COUNT OF UNIQUE TICKERS
    maxincrease = 0                     'GREATEST PERCENT INCREASE ON SHEET
    maxdecrease = 0                     'GREATEST PERCENT DECREASE ON SHEET
    maxvol = 0                          'GREATEST VOLUME ON SHEET
    max_increase_ticker = " "           'TICKERS OF GREATEST INCREASE/DECREASE/MAX VOLUME
    max_decrease_ticker = " "
    max_vol_ticker = " "
    
    ticker = Cells(2, 1).Value          'FIRST TICKER SYMBOL
    op = Cells(2, 3).Value              'OPENING PRICE OF FIRST TICKER
    cl = Cells(2, 6).Value              'CLOSING PRICE OF FIRST TICKER RECORD
    vol = 0
    
    For i = 2 To rcount
        vol = vol + Cells(i, 7).Value   'ACUMULATE TRANSACTION VOLUMES ON EACH RECORD
        If Cells(i + 1, 1).Value <> ticker Then     'INDICATES LAST RECORD FOR CURRENT TICKER
        
            If op > 0 Then                          'ONLY CALCULATE STATS IF OPENING PROCE GREATER THAN 0
                cl = Cells(i, 6).Value
                Cells(2 + tcount, 9).Value = ticker
                Cells(2 + tcount, 10).Value = cl - op
                
                If cl - op < 0 Then
                    Cells(2 + tcount, 10).Interior.ColorIndex = 3   'COLOR INDES 3 = RED
                ElseIf cl - op > 0 Then
                    Cells(2 + tcount, 10).Interior.ColorIndex = 4   'COLOR INDEX 4 = GREEN
                End If
                
                Cells(2 + tcount, 11).Value = (cl - op) / op        'CALCULATE TEARLY CHANGE AS PROPORTION OF OPENING COST
                Cells(2 + tcount, 11).NumberFormat = "0.00%"        'FORMAT YEARLY CHANGE AS A PERCENTAGE
                Cells(2 + tcount, 12).Value = vol
                
                If (cl - op) / op > maxincrease Then                'TRACK MAX PERCENT INCREASE AND DECREASE AND VOLUME
                    maxincrease = (cl - op) / op
                    max_increase_ticker = ticker
                End If
                If (cl - op) / op < maxdecrease Then
                    maxdecrease = (cl - op) / op
                    max_decrease_ticker = ticker
                End If
                If vol > maxvol Then
                    maxvol = vol
                    max_vol_ticker = ticker
                End If
                
            'IF OPENING PRICE FOR A NEW TICKER = 0, THEN CAPTURE TICKER NAME BUT LIST VALUES FOR YEARLY CHANGE, PERCENT INCREASE AND VOLUME AS 0.  LEAVE BACKGROUND COLOR FOR YEARLY CHANGE NEUTRAL.
            ElseIf op = 0 Then
                Cells(2 + tcount, 9).Value = ticker
                Cells(2 + tcount, 10).Value = 0
                Cells(2 + tcount, 11).Value = 0
                Cells(2 + tcount, 11).NumberFormat = "0.00%"        'FORMAT YEARLY CHANGE AS A PERCENTAGE
                Cells(2 + tcount, 12).Value = 0
            End If
            
            tcount = tcount + 1                 'INCREMENT COUNT OF UNIQUE TICKERS
            op = Cells(i + 1, 3).Value          'CAPTURE OPENING PRICE OF NEXT TICKER
            ticker = Cells(i + 1, 1).Value      'CAPTURE TICKER SYMBOL OF NEXT TICKER
            vol = 0                             'INITIALIZE VOLUME TO 0 FOR NEXT TICKER
           
        End If
    Next i
    
    'PRINT-OUT TABLE OF RESULTS ON EACH SHEET
    Cells(2, 15).Value = "Greatest % Increase"
    Cells(3, 15).Value = "Greatest % Decrease"
    Cells(4, 15).Value = "Greatest Total Volume"
    Cells(1, 16).Value = "Ticker"
    Cells(1, 17).Value = "Value"
    
    Cells(2, 16).Value = max_increase_ticker
    Cells(2, 17).Value = maxincrease
    Cells(2, 17).NumberFormat = "0.00%"
    
    Cells(3, 16).Value = max_decrease_ticker
    Cells(3, 17).Value = maxdecrease
    Cells(3, 17).NumberFormat = "0.00%"
    
    Cells(4, 16).Value = max_vol_ticker
    Cells(4, 17).Value = maxvol
    
Next j
MsgBox ("Workbook Analysis Complete")   'INDICATES LAST SHEET WAS COMPLETED

End Sub
