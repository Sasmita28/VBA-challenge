Attribute VB_Name = "Module1"
Sub stock_testing()

For Each ws In Worksheets

'Part1

'Declaring the variables
Dim lRow As Long
Dim lRow1 As Long
Dim ticker As String
Dim summary_row As Long
Dim yearly_change As Double
Dim percentage_change As Double
Dim total_stock As Double
Dim Greatest_increase As Double
Dim Greatest_decrease As Double
Dim Greatest_stock_volume As Double


'initializing total_stock to 0
total_stock = 0

'opening_price for the very first ticker
opening_price = ws.Cells(2, 3).Value

'assigning summary_row to 2 for summary table
summary_row = 2

'finding the last row
lRow = ws.Cells(Rows.Count, 1).End(xlUp).Row


'Headers for our new columns
ws.Cells(1, 9).Value = "Ticker"
ws.Cells(1, 10).Value = "Yearly_Change"
ws.Cells(1, 11).Value = "Percentage_Change"
ws.Cells(1, 12).Value = "Total Stock Volume"



'Looping through 2 to last row, 1st row is header

For Row = 2 To lRow
    
'assiging the 1st column values to "ticker" variable
    
    ticker = ws.Cells(Row, 1).Value
    
    
'Checking for change in ticker values
    If ws.Cells(Row + 1, 1).Value <> ws.Cells(Row, 1).Value Then
    
'if there is change in ticker value, finding the closing_price
        closing_price = ws.Cells(Row, 6).Value
        
'Assigning ticker values to summary table
        ws.Cells(summary_row, 9).Value = ticker
        
'Finding the total stock volume for each ticker
        total_stock = total_stock + ws.Cells(Row, 7).Value
        
'Finding yearly change
        yearly_change = closing_price - opening_price
        

'Formatting yearly_change column , green for positive change and red for negative change
        If yearly_change > 0 Then
        
            ws.Cells(summary_row, 10).Interior.ColorIndex = 4
            
        Else
        
            ws.Cells(summary_row, 10).Interior.ColorIndex = 3
            
        End If
        
'checking for opening price = 0 to avoid divison/0 error in percentage change.
'If 0, then we assign yearly change to 0.

        If opening_price = 0 Then
        
            percentage_change = 0
            
        Else
        
            percentage_change = yearly_change / opening_price
            
        End If
        
'Assiging new column values to respective columns
        ws.Cells(summary_row, 10).Value = yearly_change
        ws.Cells(summary_row, 11).Value = percentage_change
        ws.Cells(summary_row, 11).NumberFormat = "0.00%"
        ws.Cells(summary_row, 12).Value = total_stock
    
'incrementing summary row
        summary_row = summary_row + 1
        
 'reseting total stock to 0 , so it will count for each ticker
        total_stock = 0
        
'finding the opeining price for each ticker
        opening_price = ws.Cells(Row + 1, 3).Value
    
    Else
    
        
        
        yearly_change = closing_price - opening_price
    
        
        total_stock = total_stock + ws.Cells(Row, 7).Value
        
        
        
    End If
    


Next Row
        
     
     
'part 2

'Finding the last row for the summary table
lRow1 = ws.Cells(Rows.Count, 10).End(xlUp).Row

'declaring the variables
Dim max_percentage As Double
Dim min_percentage As Double
Dim max_stock_volume As Double

'finding max, min for percentage change and max total stock volume
max_percentage = WorksheetFunction.max(ws.Range("K2:K" & lRow1))
min_percentage = WorksheetFunction.min(ws.Range("K2:K" & lRow1))
max_stock_volume = WorksheetFunction.max(ws.Range("L2:L" & lRow1))

'assigning headers for new columns
ws.Range("P4").Value = "Greatest % Increase"
ws.Range("P5").Value = "Greatest % Decrease"
ws.Range("P6").Value = "Greatest Total Volume"
ws.Range("Q3").Value = "Ticker"
ws.Range("R3").Value = "Value"

'assiging value to the greatest % Increase
ws.Range("R4").Value = max_percentage
ws.Range("R4").NumberFormat = "0.00%"

'assiging value to the greatest % decrease
ws.Range("R5").Value = min_percentage
ws.Range("R5").NumberFormat = "0.00%"

'assiging value to the Greatest Total volume
ws.Range("R6").Value = max_stock_volume


'Finding the ticker values for greatest % increase , greatest % decrease and Greatest total volume
For i = 2 To lRow1
    If ws.Cells(i, 11) = max_percentage Then
    
        ws.Cells(4, 17).Value = ws.Cells(i, 9).Value
    
    ElseIf ws.Cells(i, 11) = min_percentage Then
        ws.Cells(5, 17).Value = ws.Cells(i, 9).Value
            
    ElseIf ws.Cells(i, 12) = max_stock_volume Then
        ws.Cells(6, 17).Value = ws.Cells(i, 9).Value
    
    End If
Next i

Next ws


End Sub
