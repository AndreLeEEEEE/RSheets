# RSheets
This program will put Rick's requests onto Monday.com, request information 

Versions of python, VS, and modules used:
- Python 3.7.8
- Openpyxl 3.0.6
- Visual Studio 2019, 16.8.6

Requirements: 
- None yet

This program's development is gonna be a bit different from the previous two since Rick's disdain of software and its like means the program should be really simple to use.
In addition, he wants to make requests with the minimum information required by the stockroom and other providers. This makes me wonder why I should be making this program if
Rick's template could just be formatted similarly to Monday.com from the get-go and then he adds Qty when he needs to and imports it into Monday.

If this program does so little, what more could I add to make its creation worthwhile? 
I can't make automatic changes to status columns since those have to be changed manually.
- When importing an excel sheet, a column can be set as a status column before and during the new board's existence if there aren't many cell values present or all values are digits
- In this new board, any values in a status column are made into labels. These labels are assigned colors from a preset order that: Red, Light Green, Yellow, and whatever other colors (I haven't tested this)
- It's not possible for a value in a status column to determine the status color
- It's not possible to make a cell color determine the status color
- It's not possible to make a value text color determine the status color
- When a pulse gets moved to another board, any columns that corresponds to a status column will have new value depending on what the old value was
1. If the old value was blank, the new value will be the recipient board's default status (in Wanco's case, it's the grey label)
2. If the old value contains something but its label color doesn't exist in the recipient board (or it doesn't have a label at all), the first label of the selection will be set
3. If the old value contains something and its label color exists in the recipient board, the color matched label will be set
What this means is that the contents of the excel cell cannot affect the label it receives during import or board move

There are three ways to carry out the goal, and one of them doesn't involve programming. However, all of them require Rick to import excel sheets to Monday.com.
1. Template with program - Rick will have a template that he'll fill out when he wants to request parts. The program filters out the non-requested parts and adds the necessary information of requests. A new excel sheet will be created with appropriate Monday.com format.
2. Template without program - The necessary information that the program adds will already be on the template, removing the need for a program.
3. No template, all program - Instead of a template with all the parts Rick could possibly need to request (which might not even be plausible), just get Rick to type in the parts he needs into command prompt and let the program make the excel sheet.

Update 12/15/2020: Alex and I have decided to just do the third method. It'll probably be easier for Rick this way since he won't have to have a template stored on his computer nor will he have to come up with a list of possible parts he could request.

Update 12/17/2020: I have sent the program executable to Rick. The only thing left to do is to wait for his feedback.

Update 12/29/2020: I have helped Rick through the process of importing excel data into Monday.com. Updated the column headers to match the new ones on the Material Request board.


