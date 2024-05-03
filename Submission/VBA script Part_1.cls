VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Sheet1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
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
            
            
         End If
   
Next i
Next

End Sub
Sub reset()
                'Declare variable Worksheet because it helps to execute in my workbook
                 Dim ws As Worksheet
                 For Each ws In Worksheets
                 ws.Range("I:L").Value = ""
                 ws.Range("O:Q").Value = ""
Next
End Sub

