Sub newtest()
' Define variable for holding the Opening Price
Dim Opening_Price As Double

 ' Define variable for holding the Closing Price
Dim Closing_Price As Double

' Define variable for holding Last Row
Dim LastRow As Long

' Define variable for holding the Ticker Name
Dim Ticker_Name As String

' Define variable for holding Yearly Change
Dim Yearly_Change As Double

' Define variable for holding Percent Change
Dim Percent_Change As Double

' Define variable for holding Total Stock Volume
Dim Total_Stock_Volume As Double

'Define variable for worksheet
Dim WorksheetName As String


' Define variables for greatest /decrease
Dim Max_Percent_Increase_Ticker As String
Dim Min_Percent_Decrease_Ticker As String
Dim Max_Total_Volume_Ticker As String
Dim Max_Percent_Increase As Double
Dim Min_Percent_Decrease As Double
Dim Max_Total_Volume As Double

' Set values for variables
Max_Percent_Increase = 0
Min_Percent_Decrease = 0
Max_Total_Volume = 0

' Keep track of the location for each ticker in the new table
  Dim Summary_Table_Row As Integer
  Summary_Table_Row = 2

' find active sheet
Dim ws As Worksheet
Set ws = ActiveSheet

' Loop through all sheets in the workbook
For Each ws In ThisWorkbook.Sheets

' begining variables
Ticker_Name = ws.Cells(2, 1).Value
Opening_Price = ws.Cells(2, 3).Value
Total_Volume = 0

' Determine the Last Row
LastRow = ws.Cells(Rows.Count, 1).End(xlUp).Row

' Loop through all tickers
For i = 2 To LastRow

' Check if we are still within the same company ticker, if it is not...
If Ticker_Name <> ws.Cells(i, 1).Value Then

' Calculate Yearly Change and Percent Change
Closing_Price = ws.Cells(i - 1, 6).Value
Yearly_Change = Closing_Price - Opening_Price

If Opening_Price <> 0 Then
Percent_Change = (Yearly_Change / Opening_Price) * 100
Else
Percent_Change = 0
End If

' Print the Ticker Name in the Summary Table
ws.Range("i" & Summary_Table_Row).Value = Ticker_Name

' Print the Yearly Change to the Summary Table
ws.Range("J" & Summary_Table_Row).Value = Yearly_Change
              
 ' Print the Percent Change in the Summary Table
ws.Range("k" & Summary_Table_Row).Value = Percent_Change
              
 'Print the Total_Stock_Volume to the Summary Table
 ws.Range("l" & Summary_Table_Row).Value = Total_Stock_Volume


' Conditional formatting for Yearly Change
If Yearly_Change < 0 Then

' red for negatives
ws.Range("J" & Summary_Table_Row).Interior.Color = RGB(255, 0, 0)
ElseIf YearlyChange > 0 Then
' green for positives changes
ws.Range("J" & Summary_Table_Row).Interior.Color = RGB(0, 255, 0)
End If


' calculate greatest increase / decrease
If Percent_Change > Max_Percent_Increase Then
Max_Percent_Increase = Percent_Change
Max_Percent_Increase_Ticker = Ticker_Name
ElseIf Percent_Change < Min_Percent_Decrease Then
Min_Percent_Decrease = Percent_Change
Min_Percent_Decrease_Ticker = Ticker_Name
End If


If Total_Stock_Volume > Max_Total_Volume Then
Max_Total_Volume = Total_Stock_Volume
Max_Total_Volume_Ticker = Ticker_Name
End If

'
Summary_Table_Row = Summary_Table_Row + 1


' Reset variables for the next Ticker
Ticker_Name = ws.Cells(i, 1).Value
Opening_Price = ws.Cells(i, 3).Value
Total_Stock_Volume = 0
End If

' Add to TotalVolume for the current Ticker
Total_Stock_Volume = Total_Stock_Volume + ws.Cells(i, 7).Value
Next i

' add greatest increase / decrease values
ws.Range("P2").Value = Max_Percent_Increase_Ticker
ws.Range("Q2").Value = Max_Percent_Increase
ws.Range("P3").Value = Min_Percent_Decrease_Ticker
ws.Range("Q3").Value = Min_Percent_Decrease
ws.Range("P4").Value = Max_Total_Volume_Ticker
ws.Range("Q4").Value = Max_Total_Volume

' Add header titles
  ws.Range("i1").Value = "Ticker"
  ws.Range("J1").Value = "Yearly Change"
  ws.Range("k1").Value = "Percent Change"
  ws.Range("l1").Value = "Total Stock Volume"
  ws.Range("O2").Value = "Greatest % Increase"
  ws.Range("O3").Value = "Greatest % Decrease"
  ws.Range("O4").Value = "Greatest Total Volume"
  ws.Range("P1").Value = "Ticker"
  ws.Range("Q1").Value = "Value"
  
' Format headers
  ws.Range("I1:l1").Font.Bold = True
  ws.Range("O2:O4").Font.Bold = True
  ws.Range("P1:Q1").Font.Bold = True


Next ws
End Sub