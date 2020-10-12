# PowerBI-ModernExcel-UserGroup
Power BI and Modern User Group Workshop


# Import Multiple Files in a folder

1. From the Power Query ribbon you click From File > From Folder:

*\Power Query Workshop\Sales_CSV

2. Select the folder "Sales_CSV"

3. Then hit the button "Combine and Transform"

4. Click "Closed and Load"

5. Insert a Table in your Power BI file (or insert a pivot table if you are in Excel) and add "Date" "Product" and "Amount" 

6. Add all the files from the Folder *\Power Query Workshop\New Sales CSV into the the folder *\Power Query Workshop\Sales_CSV

7. Click "Refresh" in the ribbon 

8. The full year 2008 is now available



# Import Multiple Sheet from one Excel Workbook


1.From the Power Query ribbon you click From File > Excel
Click on the file  in the Power Query Workshop folder "Budget.xlsx"

2. Select one of the sheet to start your transformation

3. After clicking transform you end up with the following table:

                        |Region|Market|MODEL|Date|Amount

3. Once finish you have to find a way to import the rest of the sheet

# Begginner:

4a. You duplicate the query (Right Click duplicate) and then 

5a. Click on the little gear near the step "Navigation" and select a different sheet


# Advanced:

4b. In the ribbon Home >  Advanced Editor

5b. Add the following code at the begginning : 
                        let
                        Source = (sheet_name)=>
6b. Modify this section of the code from :

        #"Budget 2500C_Sheet" = Source{[Item=sheet_name,Kind="Sheet"]}[Data],
               
               to 
                
        #"Budget 2500C_Sheet"= Source{[Item=sheet_name,Kind="Sheet"]}[Data],

  Add at the end of the Code 

          in 
              Source

7b. Click ok then rename the Queries "fCleaning_Budget_Query"



8b. Click on   New Source and Import Again the Excel file Budget.xlsx

9b. Select one sheet and click ok

10b. Delete all the steps below Source

11b. In the ribbon click Add Column > Invoke Custom Function 

12b. Select the fucntion you created previously

13b. Remove all column except "Name" and  "fCleaning_Budget_Query"

14b. Expand the Query by Clicking on the button next to the column name

15b. Change the type of the column and Rename the column accordingly

16b. Click Closed & Apply

17b. Create a table with your new data

# You are Done Congrats!!!!


