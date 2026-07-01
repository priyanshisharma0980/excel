### Q1 - To search for sales for id 112
=xlookup(112, column of 112, column of sales, "Not found", 0)     
OR    
=index(column of sales, match(112, column of 112, 0))      
OR    
=vlookup(value to search, entire table, sales column ka number, 0)     


### Extract first 2 letters from name
=left(name,2)    

### First name and last name from FULL NAME
Data - text to column - delimited   
OR    
=left(A2,search(" ", A2),-1)     
=right(A2, len(A2)-search(" ", A2))    
OR     
=textsplit(A2, " ")      


### total sales for IT department 
we have multiple rows of IT dept and various sales, we want total     
=sumif(sales column, dept column, "IT")     


### How many employees from IT and HR department
=sum(countifs(depratment column,{"IT","HR"}))    

### Protect/ Restrict ACCESS
File- Info - Protect workbook     

### Count, counta, countblank
count - only numerical values, no NULL, no alphabets     
counta - counts ALL, except blank    


### report VS dashboard
Report - static, has lots of worksheets, data in form of tables and data for in depth analysis    
Dashboard - dynamic and interactive, summarised in 1 screen, contains charts and visuals     











































