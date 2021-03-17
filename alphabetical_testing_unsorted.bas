Attribute VB_Name = "Módulo1"
Sub general_flow()

'determining how many worksheets exists in the book
Dim number_of_worksheets As Integer
number_of_worksheets = ThisWorkbook.Sheets.Count

'for loop that runs the script in every worksheet in the book
For i = 1 To number_of_worksheets
    Worksheets(i).Activate
    
    'Creating the headers of the new table
    Range("I1").Value = "Ticker"
    Range("J1").Value = "Yearly change"
    Range("K1").Value = "Percent change"
    Range("L1").Value = "Stock Volume"

    'formatting headers
    With Range("I1:L1")
            .Cells.Interior.ColorIndex = 1
            .Cells.Font.ColorIndex = 2
            .Cells.HorizontalAlignment = xlCenter
    End With
    Range("I2").Cells.HorizontalAlignment = xlLeft
    Range("L2").Cells.HorizontalAlignment = xlRight
    Range("I1:L1").Columns.AutoFit

    'calling function to analyze every ticker
    Call ticker_analysis
    
    'calling funtcion to obtain maximum increase, decrease and volume traded
    Call stats_by_ticker

Next i

End Sub

Function ticker_analysis()
'calling the function that creates the array of tickers
ticker = create_ticker_array()

'filling the columns

'variable declarations
Dim row, row_to_fill, date_min, date_max As Long
Dim volume_traded As LongLong
Dim ticker_index, ticker_size As Integer
Dim open_price, closing_price, yearly_change, percent_change As Double
row_to_fill = 2

'this loop will cycle through every ticker and retrieve all relevant data
ticker_size = UBound(ticker)
For ticker_index = 0 To ticker_size
    
    'filling ticker column
    Cells(row_to_fill, 9).Value = ticker(ticker_index)
    
    'initializing the variables
    yearly_change = 0
    closing_price = 0
    open_price = 0
    percent_change = 0
    volume_traded = 0
    date_min = 0
    date_max = 0
    
    'filling yearly change, percent change, stock volume
    'this loop will cycle every row to find all relevant data about the active ticker
    For row = 2 To 10000000
        If Cells(row, 1).Value = Empty Then Exit For 'breaks the loop when cell is empty
        If Cells(row, 1).Value = ticker(ticker_index) Then
        
            'updates the opening price if the operation happened at a previous date
            If date_min < Cells(row, 2).Value Then
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
    Cells(row_to_fill, 10).NumberFormat = "##0.00"
    If yearly_change <= 0 Then
        Cells(row_to_fill, 10).Interior.ColorIndex = 3
    Else
        Cells(row_to_fill, 10).Interior.ColorIndex = 10
    End If
    
    'filling percent change column and formatting the result to be a %
    If open_price <> 0 Then 'adding exception for opening_price =0, division by 0 is infinite
        percent_change = yearly_change / open_price
    Else
        percent_change = 0
    End If
    Cells(row_to_fill, 11).Value = percent_change
    Cells(row_to_fill, 11).NumberFormat = "0.00%"
    
    'filling stock volume and formatting so that's easier to read
    Cells(row_to_fill, 12).Value = volume_traded
    Cells(row_to_fill, 12).NumberFormat = "000,000"
    
    'formatting the table so that it's easier to read
    'formatting the table
    If row_to_fill Mod 2 = 0 Then 'determining if row is odd or even
        'formatting even rows light grey
        Cells(row_to_fill, 9).Interior.ColorIndex = 15
        Cells(row_to_fill, 11).Interior.ColorIndex = 15
        Cells(row_to_fill, 12).Interior.ColorIndex = 15
    Else
        'formatting odd rows a darker grey
        Cells(row_to_fill, 9).Interior.ColorIndex = 48
        Cells(row_to_fill, 11).Interior.ColorIndex = 48
        Cells(row_to_fill, 12).Interior.ColorIndex = 48
    End If
    
    'increases the row to fill
    row_to_fill = row_to_fill + 1

Next ticker_index

End Function


Function stats_by_ticker()
'this function will analyze each ticker in order to determine greatest increase, decrease and volume

'variable declaration
Dim max_ticker, min_ticker, volume_ticker As String
Dim max_increase, max_decrease As Double
Dim max_volume As LongLong

'initializing the variables
max_ticker = Range("I2").Value
min_ticker = Range("I2").Value
volume_ticker = Range("I2").Value
max_increase = Range("K2").Value
max_decrease = Range("K2").Value
max_volume = Range("L2").Value

'for loop to determine the correct values, one loop does it all
For i = 2 To 10000
    If Cells(i, 9).Value = Empty Then Exit For 'Breaking the loop
    If Cells(i, 11).Value > max_increase Then 'max % increase
        max_increase = Cells(i, 11).Value
        max_ticker = Cells(i, 9).Value
    End If
    If Cells(i, 11).Value < max_decrease Then 'max % decrease
        max_decrease = Cells(i, 11).Value
        min_ticker = Cells(i, 9).Value
    End If
    If Cells(i, 12).Value > max_volume Then 'max volume traded
        max_volume = Cells(i, 12).Value
        volume_ticker = Cells(i, 9).Value
    End If
Next i

'filling the data in the worksheet
'creating and formatting the table
Range("N2").Value = "Greatest % increase"
Range("N3").Value = "Greatest % decrease"
Range("N4").Value = "Greatest Total Volume"
Range("O1").Value = "Ticker"
Range("P1").Value = "Value"

Range("P4").NumberFormat = "000,000" 'formatting volume traded
Range("P2:P3").NumberFormat = "0.00%"

Range("N1:P4").Borders.ColorIndex = 1
Range("N1:P4").BorderAround ColorIndex:=1
Range("N1:P4").BorderAround Weight:=3
Range("N1:N4").Columns.AutoFit

'formatting headers
With Range("N1:P1")
    .Interior.ColorIndex = 1
    .Font.ColorIndex = 2
    .HorizontalAlignment = xlCenter
End With

'formatting titles column
With Range("N2:N4")
    .Interior.ColorIndex = 16
End With

'filling the data
Range("O2").Value = max_ticker
Range("P2").Value = max_increase
Range("O3").Value = min_ticker
Range("P3").Value = max_decrease
Range("O4").Value = volume_ticker
Range("P4").Value = max_volume
Range("P4").Columns.AutoFit
End Function

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






