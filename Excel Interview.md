### a single worksheet contains a maximum range of 1,048,576 rows (numbered 1 to 1048576) and 16,384 columns

### Absolute cell refrence and Relative cell refrence
Absolute - remains frixed when copied to other cells ($)

### Wild card
Pro* - all the words starting from pro
Pro? - only 1 chaharcter eg- prod
Pro~ searchs for * or ? eg - pro*

EG- =Countif(column, "pro*")

###  =VLOOKUP(lookup_value, table_array, col_index_num, [range_lookup])

### vlookup on ALL column in 1 go - 
apply on 1 cell and for many columns
vlookup(F2, 'score_raw'!A:GX,{2,3,4),0)

### DOUBLE VLOOKUP
1 sheet has id, name and sales
2nd sheet has name, attendance

we need to calculate attendance
=vlookupp(vlookup(A2,datatable,2,0),datatable2,2,0)

### vlookup to calculate sum
=sum(vlookup(F2, table1,{3,4,5},0)


Limitations - 
slow for large data set
it searches linearlly till it finds match
required fixed column index to retrive the data
only search right of the column
not case sensitive
need additional functions like iferror to handle in case it does not find any error


ALTERNATES are XLLOKUP and INDEX MATCH

# INDEX MATCH

index(array, row_num,column_num)
match(lookup value, array, match)
=INDEX('Score Raw'!$A:$GX,MATCH(Sheet1!$F6,'Score Raw'!$BE:$BE,0),205)
=INDEX('Score Raw'!$A:$GX,MATCH(Sheet1!F2,'Score Raw'!$BE:$BE,0),82)

### CONDITIONAL FORMATTING
HOME - conditional formatting 
To highlight duplicates

To create a new rule
#### eg - need to format the data according to value
ctrl + A the whole table

### WILDCARDS - special characters used to perform partial matches in formulas (SUMIF, VLOOKUP, COUNTIF) or searches, allowing you to find text containing specific, unknown, or variable character
3 types-
*pro - it will search any word starting from pro eg- programming
pro? - it will search only for 1 character - eg-prod
~ to include any character prod~* output - prod*

### remove all the formatting
alt +H+E+A

### NESTED IF
if(F2>90, "A",IF(F2>80,"B",IF(F2>70,"C")))

### Pivot table
to summarize analyse and represnet large state of data

### t-test in excel
add in - analysis toolpak
to compare means of 2 groups
open add in - got to t test
in A add 1 column (do not include header)
then in B add another
and give output as C
Other paranmeter as 0
it will give mean, variance and all other


### EDATE
calculate date from current date
=edate(date,months)  for years we multiply by 12

### caluculate difference between time
=TEXT(E2-D2,"hh:mm")

### AUTONUMBER the S.NO. according to name
=if(B2="","",row(1:1))

### TO calculate based on Criteria (q1 spend amount not more than 500, q2-300 etc..)
SOLVER add it
DATA - Solver
= write all the scenarios

### GOAL SEEK
if we want to achieve 5000 rs revenue
goal seek - set cell- price per unit
to value - 5000
by changing units sold

### profit = revenue - expense
### revenue - unit sold x perice per unit

### XLOOKUP

=XLOOKUP(F2,'Score Raw'!BE:BE,'Score Raw'!GW:GW)

### Add - after every 4 digit in 435893084848
with formula
=concat(left(column A,4),"-",mid(column b,5,4),"-",right(columnb,9,4))

without formula
1️⃣ Select the cell(s)
2️⃣ Press Ctrl + 1
3️⃣ Go to Number → Custom
0000-0000-0000



### What is data validation?
Data validation restricts what type of data can be entered in a cell.
If we do not want to enter integer in name column and want y and n
SELECT the rows then only - DATA - DATA VALIDATION - LIST - select source(writye y,n)
It will create a dropdown for each cell to select from (the cells that were selected)
We can use it for decimals, whole numbers, length of words

Now we want a customer error messgae if user inputs something else - 
SELECT the rows then only - DATA - DATA Validation - ERROR Alert Tab - 


### FILTER and SORT
Sort: Rearranges data order
Filter: Shows only required records
Like filter names starting with A
=FILTER('Score Raw'!$BE:$BE, LEFT('Score Raw'!BE:BE,1)="A")

# AVERAGEIFs
=IFERROR(AVERAGEIFS('Score Raw'!GX:GX, 'Score Raw'!$BE:$BE,'My Dashboard'!$C113,'Score Raw'!$CD:$CD,'My Dashboard'!E$7), "")

### COUNTIFS
=COUNTIFS('Score Raw'!$BE:$BE,'My Dashboard'!$C113,'Score Raw'!$CD:$CD,'My Dashboard'!E$7)


### How do you ensure data accuracy 
Answer: Cross-checks, control totals(before cleaning data, after cleaning same result of sum of values), conditional formatting(=A2<>B2), Data Validation, iferror

### What is FIND vs SEARCH
SEARCH is case-insensitive; FIND is case-sensitive.


### How do you optimize Excel for large datasets?
remove duplicates
remove columns that are not needed
Avoid volatile formulas
Use Tables
Use Power Query
Limit conditional formatting

### What is normalization in Excel data?
Organizing data to remove redundancy and improve consistency.


### VBA
### LOOP through rows/columns in VBA
Public Sub FunwithLoops()
Dim i As Integer
i = 1
Do While i <= 10
    ifModule
    ActiveCell.Offset(1, 0).Select     -- offset(rows,columns) this is a method  
    i = i + 1
Loop
End Sub

Offset(1, 0) = move 1 row down, same column
Offset property is used with a Range object to refer to a cell or range of cells that is a specified number of rows and columns away from a starting point. 


### USER INPUT in VBA
Public Sub funwithInputBox()
    Dim userInput As String
    userInput = InputBox("What is your Favourite color?", "Favourite color")
End Sub


### I want to AUTOMATE the sum (all columns have different range)
Public Sub automate_sum()
    Dim ws As Worksheet
    Dim lastcell As String
    For Each ws In Worksheets
        Worksheets(ws.Name).Select
        Range("F2").Select
        Selection.End(xlDown).Select
        lastcell = ActiveCell.Address(False, False)
        ActiveCell.Offset(1, 0).Select
        ActiveCell.Value = "=sum(F2: " & lastcell & ")"
    Next ws
End Sub


### Mwerge 2 columns without CONCAT
=textjoin(" ",True, A1, B1,c1)
=textjoin(delimiter, skip empty cell, columns)

### Split column first nad last name
data - convert text to column - delimited - select delimiter - text - destinaton row - click on sheet - okay

### text functions
Ms excel errors
Sorting filtering


### calculated columns are computed row-by-row during data refresh, stored in memory, and used for slicing/filtering. Measures are calculated dynamically at query time (runtime), responding to visual filters and slicers, and are more efficient for large datasets.


### IMPLICIT and EXPLICIT Measure
Implicit measures in Power BI are automatically - Eg when numerical columns alredy summition
created when dragging numeric fields into visuals, allowing for quick, simple, and non-reusable calculations. Explicit measures are manually defined by users using DAX (Data Analysis Expressions), offering full control, complex logic, reusability, and consistency across reports.






