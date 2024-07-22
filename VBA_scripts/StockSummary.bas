VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Sheet1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Sub StockDataForAllSheets()

'worksheet set up
Dim sh As Worksheet
    
'Looping through each sheet in this workbook
For Each sh In ThisWorkbook.Sheets

'Passing the current sheet to the sub
Call StockData(sh)

    Next sh
End Sub
Sub StockData(sh As Worksheet)

'Setting up variables for our analysis
    ' Variables for data processing
    Dim ticker As String
    Dim quaterlyChange As Double
    Dim percentChange As Double
    Dim totalVolume As Double
    Dim startPrice As Double
    Dim endPrice As Double
   
    ' Variables to track the stocks fluctuations
    Dim greatestPercentIncreaseTicker As String
    Dim greatestPercentIncrease As Double
    Dim greatestPercentDecreaseTicker As String
    Dim greatestPercentDecrease As Double
    Dim greatestTotalVolumeTicker As String
    Dim greatestTotalVolume As Double
    
     'Variables for iteration of each row
    Dim currentRow As Long
    Dim summaryRow As Long
    Dim lastRow As Long
   
   'Initializing variables for our analysis
   totalVolume = 0
    startPrice = 0
    endPrice = 0
    currentRow = 2
    summaryRow = 2
    lastRow = sh.Cells(sh.Rows.Count, "A").End(xlUp).Row
   
    ' Setting up the results' display
    sh.Cells(1, 9).Value = "Ticker"
    sh.Cells(1, 10).Value = "Quaterly Change"
    sh.Cells(1, 11).Value = "Percent Change"
    sh.Cells(1, 12).Value = "Total Stock Volume"
    
   
    ' Looping through each row of data
    While currentRow <= lastRow
        totalVolume = totalVolume + sh.Cells(currentRow, 7).Value
       
        ' Check if this is the last row for the current ticker
        If sh.Cells(currentRow + 1, 1).Value <> sh.Cells(currentRow, 1).Value Then
            ticker = sh.Cells(currentRow, 1).Value
            endPrice = sh.Cells(currentRow, 6).Value
           
            ' Calculating the quaterly change in stock price
            quaterlyChange = endPrice - startPrice
            
            ' Calculating the percent change in stock price
            If startPrice <> 0 Then
                percentChange = (quaterlyChange / startPrice) * 100
            Else
                percentChange = 0
            End If
           
            ' output the results to the summary table
            sh.Cells(summaryRow, 9).Value = ticker
            sh.Cells(summaryRow, 10).Value = quaterlyChange
            sh.Cells(summaryRow, 11).Value = percentChange & "%"
            sh.Cells(summaryRow, 12).Value = totalVolume
           
            ' Formatting for quaterlyChange column
            If quaterlyChange < 0 Then
                sh.Cells(summaryRow, 10).Interior.Color = RGB(255, 0, 0)
            ElseIf quaterlyChange > 0 Then
                sh.Cells(summaryRow, 10).Interior.Color = RGB(0, 255, 0)
            Else
                sh.Cells(summaryRow, 10).Interior.Color = RGB(255, 255, 255)
            End If
            
              ' Formatting for percentChange column
            If percentChange < 0 Then
                sh.Cells(summaryRow, 11).Interior.Color = RGB(255, 0, 0)
            ElseIf percentChange > 0 Then
                sh.Cells(summaryRow, 11).Interior.Color = RGB(0, 255, 0)
            Else
                sh.Cells(summaryRow, 11).Interior.Color = RGB(255, 255, 255)
            End If

           
            ' Prepare for next ticker
            summaryRow = summaryRow + 1
            totalVolume = 0
            startPrice = 0
            endPrice = 0
          ElseIf startPrice = 0 Then
        
            ' Initialize startPrice for each new ticker
            startPrice = sh.Cells(currentRow, 3).Value
        End If
       
        ' Move to the next row
        currentRow = currentRow + 1
    Wend
   
Columns("A:L").AutoFit
    
    
End Sub
