# VBA-challenge
![Wall Street real stock data analysis](resources/wall_street.jpg)

## **Background**
I was given a set of data containing every stock operation at the NYSE for 5 years and charged to write a VBA script that analyzes the data and summarizes it yearly

## **Changelog:**
**03/11/2021**
- Created the repository
- Uploaded all relevant files
- Created the script, testing pending

**03/12/2021**
- Script tested and debugged
- [ ] Need to double-check the logic for the array, try to avoid Variant declaration
- New script added, analyzes the biggest delta for stocks yearly

**03/16/2021**
- Added formatting to the readme file
- Fixed the formatting of yearly change data, new format is ##0.00 and should be easier to read
- Added formatting to the final results table
- Added formatting to ticker table, should be easier to read
- [ ] Function that cycles through every worksheet so that whole script has to run only once

**03/17/2021**
- Created new functions to clean-up general_flow() so that's easier to integrate the new function that cycles through every worksheet in the book
- Created the for loop that cycles through every worksheet in the book
- Found that a particular ticker has zeros all across the table (data corrupted maybe?), added an exception to avoid overflowing error