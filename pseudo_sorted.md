# PSEUDO CODE FOR THE SORTED DATA SOLUTION

     get initial and final date
     get opening and closing price
     set volume traded = 0
     get last row of the data table
     for loop that cycles through every row
        if date in row is < inicial date then
            update initial date
            update opening price
        end if
        if date in row is > final date then
            update final date
            update closing price
        add volume traded
        if next ticker is not the same then
            write ticker name
            calculate (closing price - opening price)
            write difference in price
            set conditional formatting for difference in price
            calculate difference in percentage
            write difference in percentage
            write volume traded
            reset dates, volume and prices
            update ticker to the next one
            increase row for the new table       
        end if
    next row
    format the results table

    get last row of the results table
    for loop that cycles every row of the results table
        get min price, max price, ticker min, ticker max, volume traded, ticker volume
        if difference price in the row < min price then
            update min price
            update ticker min
        end if
        if difference price in the row > max price then
            update max price
            update ticker max
        end if
        if volume traded in the row > volume traded then
            update ticker volume
            update volume traded
        end if
    next row
    write min price, max price, ticker min, ticker max, volume traded, ticker volume
    format the table

    for loop that will run the script on every workshhet


    
    
