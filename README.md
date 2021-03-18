# VBA-challenge
![Wall Street real stock data analysis](resources/wall_street.jpg)

## **Background**
I was given a set of data containing every stock operation at the NYSE for 3 years and commisioned to write a VBA script that analyzes the data and summarizes it yearly

## **Changelog:**
**03/11/2021**
- Created the repository
- Uploaded all relevant files
- Created the script, testing pending

**03/12/2021**
- Script tested and debugged
- [X] Need to double-check the logic for the array, try to avoid Variant declaration
- New script added, analyzes the biggest delta for stocks yearly

**03/16/2021**
- Added formatting to the readme file
- Fixed the formatting of yearly change data, new format is ##0.00 and should be easier to read
- Added formatting to the final results table
- Added formatting to ticker table, should be easier to read
- [X] Function that cycles through every worksheet so that whole script has to run only once

**03/17/2021**
- Created new functions to clean-up general_flow() so that's easier to integrate the new function that cycles through every worksheet in the book
- Created the for loop that cycles through every worksheet in the book
- Found that a particular ticker has zeros all across the table (data corrupted maybe?), added an exception to avoid overflowing error
- Successfully tested the script on test table
- Test took about 40 minutes to complete, but apparently the data is already sorted by ticker which will allow me to optimize the code and make the script run faster. Will implement a different approach
- alphabetical_testing_unsorted.* are the appropiate files for the unsorted solution
- the new sorted approach provides consistent data and the program completes much faster, I will continue developing that solution
- finished the new script, alphabetical_testing_sorted.* are the appropiate files for the sorted solution
- [X] quality control needed to guarantee that calculations are correct
- [X] implement the solution to the real data

**03/18/2021**

- reviewed the obtained data looking for possible logical mistakes, no errors were found
- implementing the script in multiple_year_stock_data.xlsx
- added a button to run the script
- script implemented on the data
- no errors found
- added multiple_year_stock_data_raw for testing purposes, MAKE A COPY and run the script on the copy
- final commit for the project