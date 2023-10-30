Attribute VB_Name = "Module1"
Sub StockAnalysis()

Dim ws As Worksheet
  
' Loop  Through All Worksheets
For Each ws In Worksheets

' Set an initial variable for holding the ticker,yearlychange,percentagechange,totalstock,openingprice,closingprice
  
Dim Ticker As String
Dim Yearly_Change As Double
Dim Percentage_Change As Double
Dim Total_Stock As LongLong
Dim Opening_Price As Double
Dim Closing_Price As Double


Yearly_Change = 0
Percentage_Change = 0
Total_Stock = 0
Opening_Price = 0
Closing_Price = 0

' Set an initial variable for holding the last row number
Dim LastRow As Long
Dim LRow_ST As Long
 
 ' Keep track of the location of stock in the summary table
Dim Summary_Table_Row As Long
Summary_Table_Row = 2
 
'set flag to determine opening_price at start of the year
Dim OpenPriceFlag As Boolean
 OpenPriceFlag = True
 
 

' Set an initial variable for holding greatest%increase,greatest%decrease,greatesttotalvolume
Dim GrtInc As Double
Dim GrtDec As Double
Dim Grtvol As LongLong

 GrtInc = 0
 GrtDec = 0
 Grtvol = 0
 

'set the column names of the summarytable

ws.Range("I1").Value = "Ticker"
ws.Range("J1").Value = "Yearly Change"
ws.Range("K1").Value = "Percentage Change"
ws.Range("L1").Value = "Total stock volume"

'set the titles for the table showing Greatest values

 ws.Range("O2").Value = "Greatest% Increase"
 ws.Range("O3").Value = "Greatest% Decrease"
 ws.Range("O4").Value = "Greatest Total Volume"
 ws.Range("P1").Value = "Ticker"
 ws.Range("Q1").Value = "Value"
 
 
 'Find the lastrow in the sheet
LastRow = ws.Cells(Rows.Count, 1).End(xlUp).Row
 
 'Loop through the stock
For i = 2 To LastRow
' Check if we are still within the same stock, if it is not...
If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then
   
   'set the stockname
       Ticker = ws.Cells(i, 1).Value
       
   'set the closing price of the stock
      Closing_Price = ws.Cells(i, 6).Value
     
    'calculate the yearly change for each stock
      
       Yearly_Change = Closing_Price - Opening_Price
      
    'calculate the percentage change
      
         If (Opening_Price = 0) Then
        Percentage_Change = 0
         Else
         Percentage_Change = Yearly_Change / Opening_Price
         End If
       
       'Calculate the totalstock
      Total_Stock = Total_Stock + ws.Cells(i, 7).Value
  
  'Print the ticker,yearlychange,percentage change,totalstock
      ws.Range("I" & Summary_Table_Row).Value = Ticker
      ws.Range("J" & Summary_Table_Row).Value = Yearly_Change
      ws.Range("K" & Summary_Table_Row).Value = Percentage_Change
      ws.Range("L" & Summary_Table_Row).Value = Total_Stock
   
   ' Apply Conditional Formatting to  Yearly_Change column
       If ws.Range("J" & Summary_Table_Row).Value >= 0 Then
        ws.Range("J" & Summary_Table_Row).Interior.ColorIndex = 10
        Else
         ws.Range("J" & Summary_Table_Row).Interior.ColorIndex = 3
       End If
    
    'Apply NumberFormat to Percentage_Change column
      ws.Range("K" & Summary_Table_Row).NumberFormat = "0.00%"
     
     
      
    'increment the summary table
      Summary_Table_Row = Summary_Table_Row + 1
   
   'Reset the variables for next stock
        Yearly_Change = 0
        Percentage_Change = 0
        Total_Stock = 0
        Opening_Price = 0
        Closing_Price = 0
  
 
    'Reset the flag
    OpenPriceFlag = True
   
   'If the cell immediately following a row is the same stock...
   
  Else
   
       'Add stock volume to get totalstock
      
      Total_Stock = Total_Stock + ws.Cells(i, 7).Value
    
      'Set the OpeningPrice of the stock
      
      If OpenPriceFlag = True Then
      Opening_Price = ws.Cells(i, 3).Value
      OpenPriceFlag = False
      End If
   

End If

Next i


'Find lastrow of summary table
  LRow_ST = Cells(Rows.Count, 9).End(xlUp).Row

   'Find Greates%inc,Greatest%dec,Greatesttotalvolume
   GrtInc = WorksheetFunction.Max(ws.Range("k2:k" & LRow_ST))
   GrtDec = WorksheetFunction.Min(ws.Range("k2:k" & LRow_ST))
   Grtvol = WorksheetFunction.Max(ws.Range("L2:L" & LRow_ST))
   
   'Print Greates%inc,Greatest%dec,Greatesttotalvolume value
     ws.Range("Q2").Value = GrtInc
     ws.Range("Q3").Value = GrtDec
     ws.Range("Q4").Value = Grtvol
            
       'Use Index and Match func to find the stocks having Greates% inc,Greatest% dec,Greatesttotalvolume
ws.Range("P2").Value = WorksheetFunction.Index(ws.Range("I2:I" & LRow_ST), WorksheetFunction.Match(GrtInc, ws.Range("k2:k" & LRow_ST), 0))
ws.Range("P3").Value = WorksheetFunction.Index(ws.Range("I2:I" & LRow_ST), WorksheetFunction.Match(GrtDec, ws.Range("k2:k" & LRow_ST), 0))
ws.Range("p4").Value = WorksheetFunction.Index(ws.Range("I2:I" & LRow_ST), WorksheetFunction.Match(Grtvol, ws.Range("L2:L" & LRow_ST), 0))
     
     
            'Apply Number Format to greatest%inc value and greatest%dec value
            ws.Range("Q2").NumberFormat = "0.00%"
            ws.Range("Q3").NumberFormat = "0.00%"
             
             'AutoFit columns
             ws.Columns.AutoFit
     
  Next ws
  
  End Sub
