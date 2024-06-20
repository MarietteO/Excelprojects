<h3>Section 13: Microsoft Excel 102 Course Introduction</h3>
<h3>Section 14: Working with an Excel List</h3>
<b>Understanding Excel List Structure</b>:<br>
The way Excel recognises your headers is by you formatting them differently.<br>
Also make sure you don't have any empty rows or columns.<br>
<b>Sorting a list using single-level sort:</b><br>
Click into the column you wish to sort -> Data Tab -> Sort Filter -> a->z and z->a to sort ascending/descending<br>
<b>Sorting a list using multi-level sorts:</b><br>
Click into any column in the list -> Data Tab -> Sort Filter -> Big Sort Button -> Opens up custom sort window. Here you can add a secondary level for a secondary sort.<br>
<b>Using custom sorts in an Excel List:(e.g., sort by month)</b><br>
Open up custom sort window (as above) -> "Order" -> "Custom List"<br>
<b>Filter an Excel List using the AutoFilter tool:</b><br>
Click anywhere in list -> Data Tab -> Sort & Filter -> Big Filter Button -> Now you will see dropdown arrows in your column headers.<br>
Note that if you had a "count" function at the bottom to count the number of records, you won't be able to see that result anymore. But you can still manually drag open that cell to see it. However, it will still show the count for the whole list, not for the fitered result.<br>
To see the whole list again, go to Sort & Filter and hit Clear<br>
<b>Creating subtotals in a list:</b><br>
First sort the column that you wish to find subtotals for -> Click into list -> Data -> Subtotal (on far right)<br>
This will also give you a grand total down at the bottom.<br>
Toggle between 1, 2, 3 at the left of your screen to see just subtotals, just grand total or everything. You can also click on the minus sign on the left to hide the details of some of the sections.<br>
<b>Formatting a list as a table (alternating row colours):</b><br>
Click anywhere in your list -> Home Tab -> Styles -> Format as Table<br>
Now you also have a new "Design" tab with features like "remove dupicates".<br>
Design -> Table Style Options -> Total Row. Now you will have a "Total" row at the bottom of your list showing the number of records, but there is a dropdown menu where you can change the function to "Sum" etc.<br>
To add new rows at the bottom (above "Total"), click on the small black triangle in the corner of the bottom right cell, and drag it down. The Total will update as you add new records.<br>
<b>Using Conditional Formatting to Find Duplicates:</b><br>
Control+Shift+Down Arrow to select the relevant column.<br>
Home -> Styles -> Conditional Formatting -> Highlight Cell Rules -> Duplicate Values<br>
<b>Removing Duplicates:</b>
If your list has been formatted as a table, you have the Design Tab -> (on the far left:)Remove Duplicates<br>
Otherwise, Data Tab -> Data Tools -> Remove Duplicates<br>
<h3>Section 15: Excel List Functions ("Database Functions")</h3>
<b>DSUM():</b> "Database sum": Summing up only a specific category in a list.<br>
For example, if you want to sum up the values for the category "Rent", found under a column header called "Category", you can, somewhere to the right of your list, type the following: Category. In the cell below that, type: Rent. Next to Category, you can type the name of the column with the numbers you wish to sum up, for instance Total Sales. Below Total Sales and right of Rent, you can now type: =DSUM() and click on the fx button. This will open up your Argument Window. For "Database", you can select your entire list. For "Field", you can type the cell number where the original "Total Sales" header resides. For "Criteria" you can select the cells to the right where you typed "Database" and "Rent". Hit OK.<br>
<b>DSUM Function with OR Criteria:</b><br>
I.e., multiple criteria under the same column (in this case, "Category"). Just type another category under "Rent", for example "Software". Go back to your formula, click on it, then click on the fx button. Adjust your "Criteria".<br>
<b>DSUM Function with AND Criteria:</b><br>
I.e., criteria from different columns. Add the name of the new column next to "Category" where you typed it on the right. For example, "Division". And underneath that, type the relevant division to consider, for example, "North". Click on your formula, click on fx, and change your criteria to include both column headers as well as the text underneath.<br>
You can also include more criteria for either column down below as needed, for example "North" (again) and "Software" (i.e., using both "AND" and "OR" statements together).<br>
<b>Other Database Functions:</b><br>
DAVERAGE(), DCOUNT(), DCOUNTA() [to count non-numeric cells]<br>
<b>Excel Function: SUBTOTAL():</b><br>
This is still a database function, i.e., something you work with inside of a vertical list.<br> tbc...

