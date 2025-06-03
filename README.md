# excel
Learning MS EXCEL

# Serial No.
1,2 then select 1,2 both and drag down.  
It will give 1,2,3,4,5,6.....  

If we selsect only 1 then all the cells will have 1,1,1,1,1.... after dragging  

# DATE  
HOME - Date option select the format)  
If you want more options for format of dates (In the date wala area in the square there is a small arrow to expand)  
Drag this date down for this format  
Step 1: Convert Text Dates to Real Dates
Select the cells with dates like 23-07-2000, 21-09-2024, etc.
Go to the Data tab → Click Text to Columns.
Choose Delimited → Click Next → Click Next again.
In Step 3, select Date and choose DMY (since your dates are in Day-Month-Year format).
Click Finish.

When there are many cells having dates in different formats and we need to set in one format select 1st one and then HOME- Format Painter and then drag down
Select FORMAT PAINTER by double clicking also

Another way - Select all the cells and right click FORMAT CELLS

# To create rows above certain rows
Right Click on 1 or A then select INSERT

Add rows ctrl + y
delete rows ctrl and -

# SUM
write =sum
then press TAB then select the cells

# Average
=average

# CONCAT
to add 2 names eg first_name + last_name = full name
In column full_name write-
=concat(e1, " ", f1)
," ", - this is used to give space


# Colour coding
give colour to rows (grren to red wala format)
go to HOME, conditional formatting, colour scales

# Filter
select the table then data - filter (to get drop down for filters)


# Duplicate
select table
then Data - remove duplicates
To highlight duplicates
Home - conditional formatting - Highlight cell rules - duplicate values


# round
=rounf(c11, 2) - it will round off to two places.
roundup() and rounddown()  

# wrap text
HOMW - Wrap text
wrap text is used to contain the whole long line in one cell

# hyperlink
select column
insert - link (ctrl+k)

# hyperlink another workbook
insert- link - place in this document

# remove hyperlink
right click the cell to edit or remove hyperlink

# drop down
create a table of these elements of dropdown
Now go to the table where you want to use this dropdown and go to
DATA - DATA validation - LIST



# SORT
# To sort 2 rows together 
select the table and add Filter, then go to sort
enter columns
Greater than, less than - select the filter for that column
Then open number filter
input the values

# SLICER 
(for example you have gender column, slicer can be moved and it will have two buttons - M and F) 
select the rows and make it a TABLe
table desgin will open in ribbon only after you make a table
select table- INSERT- SLICER - select the columns

# commnet
To add commnent, right click on the cell and insert comment

# CHARTS
Insert - recommended charts
In corner we will get + smymols and 2 more symbols
chart style,elements, filter
click on that + and click on AXIS titles to name the axis
data lables- to show data over the rectange of graph

Now that brush (chart style below + symbol)

Now below that in style we can choose filters
1. column chart - normal basic rectanfle figures


# text to column, text split
DATA - text to columns
eg- abcd@gmail.com
we need to split in two columns -abcd and gmail.com (things after @)
go to TEXT TO COLUMNS - then delimited
select others and type @ the next

# split full name into first and last name
split via a space in DATA - text to column

# Vlookup
=vlookup(lookup value, table array, column_index_number, [range_lookup])
=vlookup(name row, will select the whole refrence table, 5,0)
to lock the table - fn+f4 (2nd entry in the formula (table array)
in above example used 5 because - 5 is the position of colum salary (E) E is 5
0 is for exact match  

LIMITATION- Vlookup will ignore the duplicate values and only return of rows of the 1st value, and will ignore later duplicates


# Pivot table
to amalyse and calculate data




