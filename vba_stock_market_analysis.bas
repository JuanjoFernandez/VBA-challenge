Attribute VB_Name = "Módulo1"
Sub general_flow()

'Creating the headers of the new table
Range("I1").Value = "Ticker"
Range("J1").Value = "Yearly change"
Range("K1").Value = "Percent change"
Range("L1").Value = "Stock Volume"

'calling the function that creates the array of tickers
ticker = create_ticker_array()

'filling the columns

'variable declarations
Dim row, row_to_fill, date_min, date_max, volume_traded As Long
Dim ticker_index As Integer
Dim open_price, closing_price, yearly_change, percent_change As Double
row_to_fill = 2

'this loop will cycle through every ticker and retrieve all relevant data
For ticker_index = 0 To UBound(ticker)
    
    'filling ticker column
    Cells(row_to_fill, 9).Value = ticker(i)
    
    'filling yearly change, percent change, stock volume
    'this loop will cycle every row to find all relevant data about the active ticker
    For row = 2 To 10000000
        If Cells(row, 1).Value = Empty Then Exit For 'breaks the loop when cell is empty
        If Cells(row, 1).Value = ticker(i) Then
        
            'updates the opening price if the operation happened at a previous date
            If date_min > Cells(row, 2).Value Then
                date_min = Cells(row, 2).Value
                open_price = Cells(row, 3).Value
            End If
            
            'updates the closing price if the operation happened after the previous value
            If date_max < Cells(row, 2).Value Then
                date_max = Cells(row, 2).Value
                closing_price = Cells(row, 6).Value
            End If
            
            'updates the volume traded
            volume_traded = volume_traded + Cells(row, 7).Value
        End If
    Next row
    
    'filling yearly change column and conditional formatting the results
    yearly_change = closing_price - open_price
    Cells(row_to_fill, 10).Value = yearly_change
    If yearly_change <= 0 Then
        Cells(row_to_fill, 10).Interior.ColorIndex = 3
    Else
        Cells(row_to_fill, 10).Interior.ColorIndex = 10
    End If
    
    'filling percent change column and formatting the result to be a %
    percent_change = yearly_change / open_price
    Cells(row_to_fill, 11).Value = percent_change
    Cells(row_to_fill, 11).NumberFormat = "0.00%"
    
    'filling stock volume and formatting so that's easier to read
    Cells(row_to_fill, 12).Value = volume_traded
    Cells(row_to_fill, 12).NumberFormat = "000,000"
    
    'resetting the variables
    yearly_change = 0
    closing_price = 0
    open_price = 0
    percent_change = 0
    volume_traded = 0
    date_min = 0
    date_max = 0

Next ticker_index

End Sub

'This function creates the array that contains every individual ticker
Function create_ticker_array() As Variant

'Variable declarations
Dim ticker_array() As Variant
Dim row As Long
Dim i, i2 As Integer
Dim new_ticker As Boolean

'Since 1st row will be for sure a new ticker, this instruction places it on the array
ReDim ticker_array(0)
ticker_array(0) = Cells(2, 1).Value

'For loop that cycles through every row, the loop breaks as soon as it finds an empty cell
For row = 2 To 100000000
    
    If Cells(row, 1).Value = Empty Then Exit For 'Breaks the cycle if cell is empty
        
    'Checks the array for the value found, and adds it to the array if it finds a new ticker
    i2 = UBound(ticker_array)
    For i = 0 To i2
        new_ticker = True
        If Cells(row, 1).Value = ticker_array(i) Then
            new_ticker = False
            Exit For
        End If
    Next i
        
    If new_ticker = True Then
        ReDim Preserve ticker_array(UBound(ticker_array) + 1)
        ticker_array(i) = Cells(row, 1).Value
    End If

Next row

'This block tests if the array is created correctly, will be commented after a successful test
'Dim array_test As String
'For i = 0 To UBound(ticker_array)
'    array_test = array_test + " " + ticker_array(i)
'Next i
'MsgBox "Los contenidos del arreglo son:" + array_test

'Finishing the function and declaring the returning value
create_ticker_array = ticker_array

End Function






