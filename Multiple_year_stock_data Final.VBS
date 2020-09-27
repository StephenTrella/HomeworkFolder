Attribute VB_Name = "Module1"
Sub Stocks()
Dim ws As Worksheet
   For Each ws In Sheets
   ws.Activate
Cells.Select
Cells.EntireColumn.AutoFit
 
 'Defining headers'
Range("I1").Value = "Ticker"
Range("J1").Value = "Yearly Chan"
Range("K1").Value = "Percent Chan%"
Range("L1").Value = "Stock"
Range("O2").Value = "Greatest % Increase"
Range("O3").Value = "Greatest % Decrease"
Range("O4").Value = "Greatest Total Volume"
Range("P1").Value = "Ticker"
Range("Q1").Value = "Value"

'Defining Variables'
Dim Ticker As String
Dim i As Long
Dim FirstValue As Double
Dim LastValue As Double
Dim Diff As Double
Dim FirstVol As Double
Dim EndVol As Double
Dim TotalVol As Double
Dim YrChng As Double
Dim YrOpen As Double
Dim PerChng As Double
Dim LastRow As Long

YearlyOpen = Cells(2, 3).Value
TotalVol = 0
Summary_table_row = 2
LastRow = (Cells(Rows.Count, 1).End(xlUp).Row)

'For Loop Statement'
 For i = 2 To LastRow
   If Cells(i + 1, 1).Value <> Cells(i, 1).Value Then
      Ticker = Cells(i, 1).Value

'Calculate Ticker/Total Vol/Yearly Change'
      TotalVol = TotalVol + Cells(i, 7).Value
      
           Range("I" & Summary_table_row).Value = Ticker
           Range("L" & Summary_table_row).Value = TotalVol
      YearlyClose = Cells(i, 6).Value
      YearlyChange = YearlyClose - YearlyOpen
                  Range("J" & Summary_table_row).Value = YearlyChange
      TotalVol = 0
           'Calculate Percent Change'
                 If (YearlyOpen = 0 And YearlyClose = 0) Then
                  PercentChange = Null
                  
                  ElseIf (YearlyOpen = 0 And YearlyClose <> 0) Then
                  PercentChange = -1
                  
                  Else
                  PercentChange = (YearlyChange / YearlyOpen)
                  ws.Range("K" & Summary_table_row).Value = PercentChange
                  ws.Range("K" & Summary_table_row).NumberFormat = "0.00%"
                  
                  End If
                  Summary_table_row = Summary_table_row + 1
                  TotalVolume = 0
                  YearlyOpen = Cells(i + 1, 3)
  Else
  TotalVol = TotalVol + ws.Cells(i, 7).Value
  
  End If
  
'Conditional formatting that will highlight positive change in green and negative change in red'
If Range("J" & Summary_table_row).Value >= 0 Then

      Range("J" & Summary_table_row).Interior.ColorIndex = 4
  Else
      Range("J" & Summary_table_row).Interior.ColorIndex = 3
  End If
Next i
 
 'Find "Greatest % increase", "Greatest % decrease" and "Greatest total volume"
Dim Greatest As Double
Dim Decrease As Double
Dim GreatestVol As Double

Greatest = Application.WorksheetFunction.Max(ws.Range("K2:K" & Summary_table_row))
Decrease = Application.WorksheetFunction.Min(ws.Range("K2:K" & Summary_table_row))
GreatestVol = Application.WorksheetFunction.Max(ws.Range("L2:L" & Summary_table_row))

For J = 2 To Summary_table_row
 If Cells(J, 11).Value = Greatest Then
   Cells(2, 16).Value = Cells(J, 9).Value
   Cells(2, 17).Value = Cells(J, 11).Value
   Cells(2, 17).NumberFormat = "0.00%"
ElseIf Cells(J, 11).Value = Decrease Then
  Cells(3, 16).Value = Cells(J, 9).Value
   Cells(3, 17).Value = Cells(J, 11).Value
   Cells(3, 17).NumberFormat = "0.00%"
ElseIf Cells(J, 12).Value = GreatestVol Then
    Cells(4, 16).Value = Cells(J, 9).Value
    Cells(4, 17).Value = Cells(J, 12).Value
End If
Next J
Next ws
End Sub


