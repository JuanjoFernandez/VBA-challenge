Attribute VB_Name = "Módulo1"
Sub general_flow()
'for loop that cycles through every worksheet in the workbook
Dim ws As Worksheet
For Each ws In Sheets
    
    'making current iteration the active sheet
    ws.Activate
    
    'calling the function that provides the results table
    Call data_loop

    'calling the function that gives max values for the results table
    Call results_loop

    'calling the function to format the table
    Call worksheet_format

Next ws

End Sub
'This function loops through all the data in the table and retrieves all relevant information
Function data_loop()

'variable declaration
Dim ticker_name As String
Dim date_min, date_max As Long
Dim open_price, close_price As Double
Dim volume_traded As LongLong
Dim last_row, row_to_fill As Long

'variable initialization
ticker_name = Range("A2").Value
date_min = Range("B2").Value
date_max = Range("B2").Value
open_price = Range("C2").Value
close_price = Range("F2").Value
volume_traded = 0
row_to_fill = 2
'getting the last row with data
last_row = Cells(Rows.Count, 1).End(xlUp).row

'creating the headers for the new table
Range("I1").Value = "Ticker"
Range("J1").Value = "Yearly change"
Range("K1").Value = "Percent change"
Range("L1").Value = "Total Stock Volume"

'for loop that will cycle through every row of the data table
For i = 2 To last_row
    
    'evaluating the dates
    If Cells(i, 2).Value < date_min Then
        date_min = Cells(i, 2).Value
        open_price = Cells(i, 3).Value
    End If
    If Cells(i, 2).Value > date_max Then
        date_max = Cells(i, 2).Value
        close_price = Cells(i, 6).Value
    End If
    
    'updating the volume traded
    volume_traded = volume_traded + Cells(i, 7).Value
    
    'checking if next ticker is a new one
    If ticker_name <> Cells(i + 1, 1).Value Then
    
        'filling the results table
        'filling and formatting the ticker name
        Cells(row_to_fill, 9).Value = ticker_name
        Cells(row_to_fill, 9).Font.FontStyle = Bold
        
        'calculating, filling, and conditional formatting price difference
        Cells(row_to_fill, 10).Value = close_price - open_price
        If Cells(row_to_fill, 10).Value > 0 Then
            Cells(row_to_fill, 10).Interior.ColorIndex = 4
        Else
            Cells(row_to_fill, 10).Interior.ColorIndex = 3
        End If
        Cells(row_to_fill, 10).NumberFormat = "#0.00"
        
        'calculating, filling and formatting percentage difference
        'preventing overflow caused by division by zero
        If open_price = 0 Then
            Cells(row_to_fill, 11).Value = 0
        Else
            Cells(row_to_fill, 11).Value = (close_price - open_price) / open_price
        End If
        Cells(row_to_fill, 11).NumberFormat = "0.00%"
        
        'filling and formatting volume traded
        Cells(row_to_fill, 12).Value = volume_traded
        Cells(row_to_fill, 12).NumberFormat = "000,000"
        
        'updating all variables
        ticker_name = Cells(i + 1, 1).Value
        date_min = Cells(i + 1, 2).Value
        date_max = Cells(i + 1, 2).Value
        open_price = Cells(i + 1, 3).Value
        close_price = Cells(i + 1, 6).Value
        volume_traded = 0
        
        'moving to the next row in the results table
        row_to_fill = row_to_fill + 1
    End If
    
Next i
End Function

'this function loops through the results table to find the greatest increase, decrease and volume traded
Function results_loop()

'variable declaration
Dim increase_ticker, decrease_ticker, volume_ticker As String
Dim increase_value, decrease_value As Double
Dim volume_value As LongLong
Dim last_row As Long

'variable initialization
increase_ticker = Range("I2").Value
decrease_ticker = Range("I2").Value
volume_ticker = Range("I2").Value
increase_value = Range("K2").Value
decrease_value = Range("K2").Value
volume_value = Range("L2").Value

'getting the last row of the results table
last_row = Cells(Rows.Count, 9).End(xlUp).row

'for loop that cycles through every row of the results table
For i = 2 To last_row
    
    'Comparing to find if % increase is greater
    If Cells(i, 11).Value > increase_value Then
        increase_ticker = Cells(i, 9).Value
        increase_value = Cells(i, 11).Value
    End If
    
    'Comparing to find if % decreased is greater
    If Cells(i, 11).Value < decrease_value Then
        decrease_ticker = Cells(i, 9).Value
        decrease_value = Cells(i, 11).Value
    End If
    
    'Comparing to find if volume traded is greater
    If Cells(i, 12).Value > volume_value Then
        volume_ticker = Cells(i, 9).Value
        volume_value = Cells(i, 12).Value
    End If
    
Next i

'building the new table

'headers for rows and columns
Range("O1").Value = "Ticker"
Range("P1").Value = "Value"
Range("N2").Value = "Greatest % increase"
Range("N3").Value = "Greatest % decrease"
Range("N4").Value = "Greatest Total Volume"

'filling in the values found
Range("O2").Value = increase_ticker
Range("O3").Value = decrease_ticker
Range("O4").Value = volume_ticker
Range("P2").Value = increase_value
Range("P3").Value = decrease_value
Range("P4").Value = volume_value

'formatting the values
Range("P2").NumberFormat = "0.00%"
Range("P3").NumberFormat = "0.00%"
Range("P4").NumberFormat = "000,000"
End Function

'this function formats all the results
Function worksheet_format()

'formatting results table first

'formatting headers
With Range("I1:L1")
    .HorizontalAlignment = xlCenter
    .Interior.ColorIndex = 1
    .Font.ColorIndex = 2
    .Font.FontStyle = "Bold"
End With

'color formatting rows so they are easier to read
'getting last row
Dim last_row As Long
last_row = Cells(Rows.Count, 9).End(xlUp).row

'for loop to color format and apply text allignment the rows
For i = 2 To last_row
    Cells(i, 9).HorizontalAlignment = xlLeft
    If i Mod 2 = 0 Then
        Cells(i, 9).Interior.ColorIndex = 48
        Cells(i, 11).Interior.ColorIndex = 48
        Cells(i, 12).Interior.ColorIndex = 48
    Else
        Cells(i, 9).Interior.ColorIndex = 15
        Cells(i, 11).Interior.ColorIndex = 15
        Cells(i, 12).Interior.ColorIndex = 15
    End If
Next i

'autofitting columns
Range("I:P").Columns.AutoFit

'formatting the final table
With Range("N1:P1")
    .Interior.ColorIndex = 1
    .Font.ColorIndex = 2
    .Font.FontStyle = "Bold"
End With
    
Range("N2:N4").Interior.ColorIndex = 15
Range("N1:P4").Borders.ColorIndex = 1
Range("N1:P4").BorderAround ColorIndex:=1, Weight:=xlThick
End Function
