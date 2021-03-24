Sub Stoks_A()


Dim ws As Worksheet

For Each ws In Worksheets
ws.Activate

'variables para type of stocks
Dim Ticker_name As String
Dim Ticker_row As Integer
Dim i, ult As Long
Dim Percent_Change, Yearly_Change, Yearly_Opening, Yearly_Close As Double
Dim Total_Stock_Volume, Greatest_Total_Volume As LongLong
Dim Greatest_Percentage_Increase, Greatest_Percentage_Decrease, Last_Row As Integer

'----------------CREATE THE TABLE 1
ws.Range("J1,P1") = "Ticker"
ws.Range("K1") = "Yearly_Change"
ws.Range("L1") = "Percent_Change"
ws.Range("M1") = "Total_Stock_Vol"
ws.Range("Q1") = "Value"

'----------------Ultima celda de la colunma-------------------
'set values for variables

Ticker_row = 2

Yearly_Change = 0

Yearly_Opening = ws.Cells(2, 3).Value

Yearly_Close = 0

Total_Stock_Volume = 0

Ticker_name = ws.Cells(2, 1).Value


'define the last row of Ticker
ult = Cells(Rows.Count, 1).End(xlUp).Row
'--------------------Ticker loop

  For i = 2 To ult
'-------check if we have reached the last row of that Ticker
  If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then

'--------------------'set the ticker_name
   Ticker_name = ws.Cells(i, 1).Value
     
 '--------------------Total_Stock_Volume
 
  Total_Stock_Volume = Total_Stock_Volume + ws.Cells(i, 7).Value

'--------------------Yearly_Close Price

   Yearly_Close = ws.Cells(i, 6).Value
   
'--------------------Yearly_Change of Price
   
   Yearly_Change = (Yearly_Close - Yearly_Opening)
   
   'Check ----------------------Percent_Change

     If Yearly_Opening = 0 Then
         
         Percent_Change = 0
     
     Else
        
        Percent_Change = (Yearly_Close - Yearly_Opening) / Yearly_Opening
     
     End If
     
    '--------------------SAVING IN TABLE !
'save ticker_name in the table
   ws.Range("J" & Ticker_row).Value = Ticker_name
'save Yearly_Change in the table
   ws.Range("K" & Ticker_row).Value = Yearly_Change
 'save Percent_Change in the table
   ws.Range("L" & Ticker_row).Value = Percent_Change
   ws.Range("L:L").NumberFormat = "0.00%"
'save Total_Stock_Volume in the table
   ws.Range("M" & Ticker_row).Value = Total_Stock_Volume
   
       ws.Range("J1:Q1").Columns.AutoFit
       ws.Range("O2:O4").Columns.AutoFit
       ws.Cells(2, 15).Value = "Greatest_Percentage_Increase"
       ws.Cells(3, 15).Value = "Greatest_Percentage_Decrease"
       ws.Range("Q2:Q3").NumberFormat = "0.00%"
       ws.Cells(4, 15).Value = "Greatest_Total_Volume"
       ws.Range("J1:Q1").HorizontalAlignment = xlCenter
 
   
'sum 1 to rows to make it go down 1 on table of totals
   Ticker_row = Ticker_row + 1
 
 'reset Total for the following Ticker
    Total_Stock_Volume = 0
    Percent_Change = 0
    Total_Stock_Volume = 0
    
 Else
 
 'adding the Total_Stock Volume
Total_Stock_Volume = Total_Stock_Volume + ws.Cells(i, 7).Value

 If ws.Cells(i, 11).Value > 0 Then
    ws.Cells(i, 11).Interior.ColorIndex = 4
  ElseIf ws.Cells(i, 11).Value < 0 Then
      ws.Cells(i, 11).Interior.ColorIndex = 3
  End If
  
 End If
'---------------- Formatting
  If ws.Cells(i, 11).Value > 0 Then
    ws.Cells(i, 11).Interior.ColorIndex = 4
  ElseIf ws.Cells(i, 11).Value < 0 Then
      ws.Cells(i, 11).Interior.ColorIndex = 3
  End If
       
      
     Next i
   Next ws

End Sub

            
            
            
            
            
   

