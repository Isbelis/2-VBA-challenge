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

Sub stocks_data():

' Declare the variables:
   ' Declare ticker
      Dim ticker As String

  ' Declare Quarterly price open
     Dim ticker_open As Double
   
  'Declare  Quarterly price close
     Dim ticker_close As Double

   ' Declare Quarterly change
      Dim price_change As Double
   
   ' Declare total stock volume column G and L
     Dim cell_vol As LongLong
     Dim vol_total As LongLong
   ' Declare ercentage change
      Dim percent_change As Double

    ' Declare to set up a row to start
      Dim k As Long ' leaderboard row
      
    'Declare variable Worksheet because it helps to execute in my workbook
    ' Declare a variable to store the active worksheet
     Dim ws As Worksheet
    
           
'**********************************************
' Loop in all worksheet
'**********************************************
  ' Loop through each worksheet in the workbook
     For Each ws In ThisWorkbook.Worksheets

' Insert label of my new Columns
    ws.Range("I1").Value = "Ticker"
    ws.Range("J1").Value = "Quarterly Change"
    ws.Range("K1").Value = "Percent Change"
    ws.Range("L1").Value = "total stock volume"
    ' Def to star
    k = 2
    previous_i = 1
    vol_total = 0

   ' asked the xpert
    Dim lastRow As Long
    lastRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
    Dim i As Long ' row number
     ' Loop the quarterly  change, % change, and total stock volume by each ticker name
      For i = 2 To lastRow
      cell_vol = ws.Cells(i, 7).Value
      ' LOOP rows 2 to 254
        ' check if the next row's ticker is DIFFERENT from the current one
        ' if the same, then we only need to add to the vol_total
        ' if DIFFERENT, then we need add last row, write out to the leaderboard
        ' reset the vol_total to 0
      If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then
      
    ' Get the Ticker variable
      ticker = ws.Cells(i, 1).Value
       
             'Intiate the variable to go to the next Ticker Alphabet

            previous_i = previous_i + 1

            ' Get the value first day open form the column 3 or "C" and last day close of the year on column 6 or "F"

            ticker_open = ws.Cells(previous_i, 3).Value
            ticker_close = ws.Cells(i, 6).Value

            ' A for loop to sum the total stock volume using vol which is found in column 7 or "G"

            For j = previous_i To i

               ' OMG we have a different ticker, panic
            vol_total = vol_total + cell_vol
            Next j
    
        End If
        
     Next i
 Next
End Sub

