VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Sheet1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Option Explicit
'*******************************************************
' My VBA script file are three:
' Level 1 : Define variables according numeral from 1 and 4 from my readme.
' Level 2 : Include Level and Using conditional formatting according to point 5 of my readme.
' Level 3 : Include Level 1 and 2, additional I determined the stock with the "Greatest % increase", "Greatest % decrease", and "Greatest total volume".
'**************************************************

Option Explicit

Sub stocks():
' Declare the variables:

    'Declare variable Worksheet because it helps to execute in my workbook
     Dim ws As Worksheet
     
     For Each ws In Worksheets
     

    Dim i As Long ' row number
    Dim cell_vol As LongLong ' contents of column G (cell volume)
    Dim vol_total As LongLong ' what is going to go in column L
    Dim ticker As String ' what is going to go in column I

    Dim k As Long ' leaderboard row
    
    Dim ticker_close As Double
    Dim ticker_open As Double
    Dim price_change As Double
    Dim percent_change As Double
    percent_change = 0
    
    ' asked the xpert
    Dim lastRow As Long
    lastRow = ws.Cells(Rows.Count, 1).End(xlUp).Row

    vol_total = 0
    k = 2
    
    ' Write Leaderboard Columns
    ws.Range("I1").Value = "Ticker"
    ws.Range("J1").Value = "Quarterly Change"
    ws.Range("K1").Value = "Percent Change"
    ws.Range("L1").Value = "Volume"
    
    
    ' assign open for first ticker
    ticker_open = Cells(2, 3).Value

    For i = 2 To lastRow:
        cell_vol = Cells(i, 7).Value
        ticker = Cells(i, 1).Value

        ' LOOP rows 2 to 254
        ' check if next row ticker is DIFFERENT
        ' if the same, then we only need to add to the vol_total
        ' if DIFFERENT, then we need add last row, write out to the leaderboard
        ' reset the vol_total to 0

        If (Cells(i + 1, 1).Value <> ticker) Then
            ' OMG we have a different ticker, panic
            vol_total = vol_total + cell_vol
            
            ' get the closing price of the ticker
            ticker_close = Cells(i, 6).Value
            price_change = ticker_close - ticker_open
            
            ' Check if open price is 0 first - to protect from Kevin the Hacker
            If (ticker_open > 0) Then
                percent_change = price_change / ticker_open
        
                ticker_open = ws.Cells(i + 1, 3).Value
             End If

            ws.Cells(k, 9).Value = ticker
            ws.Cells(k, 10).Value = price_change
            ws.Cells(k, 11).Value = percent_change
            ws.Cells(k, 12).Value = vol_total
            
            ' formatting
            If (price_change > 0) Then
                ws.Cells(k, 10).Interior.ColorIndex = 4 ' Green
            ElseIf (price_change < 0) Then
                ws.Cells(k, 10).Interior.ColorIndex = 3 ' Red
            Else
                ws.Cells(k, 10).Interior.ColorIndex = 2 ' White
            End If
         

            ' reset
            vol_total = 0
            k = k + 1
         
      
        Else
            ' we just add to the total
            vol_total = vol_total + cell_vol
        End If
    Next i
    
    ' Style my leaderboard
    ws.Columns("K:K").NumberFormat = "0.00%"
    ws.Columns("I:L").AutoFit
   
      'Highlight the stock price changes
    'First label the cells according to the sample .png provided in the assignment
  
        ws.Cells(2, 15).Value = "Greatest % Increase"
        ws.Cells(3, 15).Value = "Greatest % Decrease"
        ws.Cells(4, 15).Value = "Greatest Total Volume"
        ws.Cells(1, 16).Value = "Ticker"
        ws.Cells(1, 17).Value = "Value"


          'Determine the max and min values in column "Percent Change" and just max in column "Total Stock Volume"
    'Then collect the ticker name, and the corresponding values for the percent change and total volume of trade for that ticker
    
    ' Declare variables to define increase

                      Dim MaxValue As Integer
                      MaxValue = 0
                      Dim MinValue As Integer
                      MinValue = 0
                      Dim Great_Volum As Integer
   
    
        For k = 2 To lastRow
            'Find the maximum percent change
            
           
            
            ' Find Max percentage change
               ws.Range("Q2") = "%" & WorksheetFunction.Max(Range("K2:K" & lastRow)) * 100
               
               ' Find Min percentage change
               ws.Range("Q3") = "%" & WorksheetFunction.Min(Range("K2:K" & lastRow)) * 100
               
             'Ticker on I Ticker
             
                 MaxValue = WorksheetFunction.Match(WorksheetFunction.Max(Range("K2:K" & lastRow)), Range("K2:K" & lastRow), 0)

                 ws.Range("P2") = ws.Cells(MaxValue + 1, 9)
                 
                 MinValue = WorksheetFunction.Match(WorksheetFunction.Min(Range("K2:K" & lastRow)), Range("K2:K" & lastRow), 0)

                 ws.Range("P3") = ws.Cells(MinValue + 1, 9)
                 
                 'Find Greatest Total volume
                 
                 Great_Volum = WorksheetFunction.Match(WorksheetFunction.Max(Range("L2:L" & lastRow)), Range("L2:L" & lastRow), 0)

                   ws.Range("P4") = ws.Cells(Great_Volum + 1, 9)
                   ws.Range("Q4") = WorksheetFunction.Max(Range("L2:L" & lastRow))
                   
                   
                   
                   
                   
                   
                   
   


   Next k
   Next
        
End Sub

Sub reset()
                'Declare variable Worksheet because it helps to execute in my workbook
                 Dim ws As Worksheet
                 For Each ws In Worksheets
                 ws.Range("I:L").Value = ""
                 ws.Range("I:L").Interior.ColorIndex = 2
                 ws.Range("O:Q").Value = ""
Next
End Sub
