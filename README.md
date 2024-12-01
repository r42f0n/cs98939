java c
Advanced Excel: Excel Fundamentals review
Module 1 Lecture
Module 1   – Excel interface and basic calculations
Outline
Part 1 Excel interface and   formatting
•         Visiting   the Excel Interface
-            Overview of   the Excel interface
-            Navigating ribbons and panels
-            Formula bar, Name Box, and Formula editor
-            Insert   function and Editing area
-            Status bar and   Zoom
-          Automatic calculator   for selected area
•         Structure of   the Excel Document
-            Renaming and organizing   worksheets
-            Maximum number of columns and rows
-          Cell content   types and limitations
-          Creating, saving, and opening   workbooks
-            Deleting and reorganizing   worksheets
•          Navigating   the   Workbook and Selecting
-            Keyboard shortcuts   for navigation
-          Selecting regions and specific cells
-            Formatting and deleting cells
-          (optional)   Accessing ribbon options   with   the keyboard
•          Managing Columns and Rows
-            Clearing contents and   formats
-            Inserting and deleting columns and rows
-            Hiding and unhiding columns and rows
-            Resizing columns and rows
•          Formatting Cells
-            Changing display   formats
-          Aligning   text and adding borders
-            Merging cells and   wrapping   text
•          Building Charts
-            Creating and customizing charts
-            Moving charts   to new   worksheets
-            Customizing chart elements
-          Adding and   formatting data labels
Part 2 Formulas and Basic   Calculations
•          Creating Formulas   to Perform. Calculations
-          Typing   formulas and using operators
-            Using cell references
-            Copying and reusing   formulas
-          Absolute   vs. relative cell references
-          Linking cells between   worksheets and   workbooks
•          Naming Cells and Ranges
-          Using predefined   functions   for calculations
-            Naming and using named ranges in   formulas
Part 3 Formulas   with Basic   function and   Text Functions
•          Basics of Predefined Functions
-            Sum,   Average, Product, Min, Max
-          Text   functions (Len, Find,   Trim, etc.)
-            Rounding   functions (Round, Int,   Trunc)
-          Counting   functions (Count, CountA, CountBlank)
Part 3 Formulas   with Basic Date and   Time Functions
•          Date and   Time in   Excel
-          Storing and using date and   time information
-            Date and   time   functions
Part 1. Excel   Interface   Visiting   the Excel interface
   
o   Caption bar:   AutoSave option (Office 365)
o   Ribbons navigate   to different panels
            File button: File > Options > Customize Ribbon
            Formulas
o   Row: 1:1048576; Columns:A:   XFD
o   Name Box: shows   the coordinate of   the cell/ find a   far cell, e.g.,   Z3000/rename a cell,   20:00
o   Insert   function button:   function library (>3000   functions in Excel)
o   Formula bar: display   the formula used in a cell
o   Worksheet   Tab
o   Zoom in/ out
o   Status Bar:   Automatic calculator   for   the selected area, customize   the results shown
File: Module1 - Part1 Excel interface   and   formatting.xlsx
Task 1: Create a new   workbook
o   Method 1:
File > New > Select a Blank   workbook or from existing   templates
Task 2: Save a   workbook
o   AutoSave: may not always be activated, depending on   where   you save it.   Note: it   will also overwrite   the previous document.
o   File > Save (using the default name and   format)
o   File > Save   As > Customize the name/ format/ location.
Task 3: Open an existing document
o   File > Open from OneDrive (saved in   the cloud)
o   File > Browse,   to select a local file
Structure of   the Excel document
An excel   file is a Collection of   Worksheets in   the opened   Workbook
Note:   the number of   worksheets is limited by   the computers memory the   worksheets default name   depends on   the language   youre using in   your Excel
Task   4: Rename   worksheet
o   right-click   then select rename, or double-click   Note:
-          The   worksheet name cannot be empty, better   to have a meaningful name, but not   too long   (<32 characters).
-          You cant use any of   the   following characters in   the   worksheets name: \/*[]:?
-          Try not   to use:   `~!@#$%^*()-space=+{}|;:,<.>
Task 5: Create/Delete a   worksheet
o   Press   the + button near   the   worksheet   tabs
o   Right-click   the   worksheet    Remove   Task 6: Re-organize the   worksheet
o   Right-click   the   worksheet    Move or Copy   Task 7: Hide and Unhide a   worksheet
o   Right-click   the   worksheet    Hide
Columns:
-            Max number of columns: 16384
-          Indexed by   letters from   A   to   Z.   The 27th    column is   AA.   The last column is   XFD,   the   16384th       column. (Ctrl + Right)
Rows:
-            Indexed by numbers, 1048576. (Ctrl + Down)
Cells
-          An   Active cell: click on a cell,   you may edit   the content in it.
-            In total, maximum 17,179,869,184 cells ~ 17 millions
-          Text: Max size   for a   text   within a cell is 32767 chars
-            Number:
-            Dates: bw 01/01/1900 and 31/12/9999
-            Time: 00:00:00 and 9999:59:59
-            A Boolean   value:   True/ False (Vrai/Faux)
-             Formula: < 8192 chars (¼ the max text size)
Note:
-            Excel automatically detects   the content in   the cell.   Try:   This is a   value, 12, 12.34, 12,34
-             English Excel recognizes 12.34 as a number, but not 12,34
-          No   worry if   you open an English Excel   file from French Excel,   the   translation   will   be   automatically done.
-          Try:   true,   false,   faux, 10:59, 10:59:25, 01/01/2020, 1/3, 15/08/1769   Task: Navigate using the direction keys
o   Ctrl + direction keys
o   Ctrl + Enter (fulfill   the selected region)
Navigating   the   workbook and selecting
-            Keyboard shortcuts   for navigation
-          Selecting regions and specific cells
-            Formatting and deleting cells
-          (optional)   Accessing ribbon options   with   the keyboard
The best Excel users always use   the keyboard!
Keyboard shortcuts: MS support [link], shortcuts for both   Windows and MacOS.   Navigating the document:
•          Move to   the next cell:   Tab (horizontally) or Enter (vertically)
•          Move   to   the last cell containing a   value: Ctrl + Down
•          Select a region :Select an edge cell in   the middle, so   you see   the fat   white cross,   then select   until   the diagonal cell.
White cell means an active cell, once   you enter a   value, it   will receive   that   value.   Click Enter   to move   within   the selected region.
•          Overwrite the selected region by   the active cell   value: Ctrl + Enter
•          Go back before   the last modification: Ctrl +   Z
•          Select non-adjacent cells: select   the   first region, press Ctrl, select   the next region.   Note: do not include a single cell if   you dont   want!
•          Select adjacent cells   with   values: Ctrl + Shift + direction keys
•          Select   the   value block: Ctrl +   A
(twice   will select   the entire   worksheet)
•          Select   the entire   worksheet:
o   Click   twice Ctrl +   A
o   Click   the gray   triangle at   the top left corner
•          Select   one/multiple   columns:
o   Click   the column name, e.g., H, see   the mouse turns to a black down arrow,   then   move to   the   end.
o   Try   with Ctrl and   with Shift
•          Change   font   to italic: Ctrl + I
•          Change   font   to underline: Ctrl + U
•          Change   font   to bold: Ctrl + B
•          Clear content: Select   then press the Delete key
•            Delete cells:
o   Select the region > right-click > Delete
o   Select the region > Ribbon Editing > Clear
Note: Right-click   will give   you different options depending on   your selected object.
•          Access to   the ribbon   with the keyboard: Press   Alt button
o   Alt H      E   C
Managing Columns and rows
•            Insert columns:
o   Insert a new column (multiple columns)   to   the left:   Select   the column> Right-click   the   column label > Insert
•          Hide columns: Select   the column > Right-click >Hide
•            Unhide columns:
o   Select the columns before and after the hidden ones > Right-click > Unhide
•            Hide and unhide rows
•            Re-size columns:
o   Manually adjust: put   the mouse at   the edge of   the column,   when it shifts to a   vertical   line   with   two arrows (flying mouse) directing both directions, right-click and drag
o   With given measurement: select   the column/columns > right-click > Column   Width >   insert   the pre-defined measurement
o   Autofit depending on   the content: put   the mouse at   the edge of   the column,   when it   shifts   to a   flying mouse >   Double-click
o   Autofit multiple columns simultaneously: select columns > double-click any edge of   the selected columns
•          Re-size rows:   work   the same as above
Note:   A hashtag sign   # appears   when   the column   width is   too narrow   to display   the   value
Formatting   the cells
The cell contains   two   types of information:   the   value or   function, and   the display   format.
•          Change display   format
o   Ribbon Home
o   Select the region > Right-click > Format Cells
Note: it does not mean   the   value has been changed. (Check   the Formula Bar.)
•         Alignment
•            Font
•            Border
Note: that even in the   worksheet display,   the gray grid is   visible, however it   will not be printed.   We   need here   to add borders to   the printed   table.   (Try File Print, see default result)
•          Wrap   Text
•            Merge Cells and center
Task 1: Display numbers   with no decimal
o   Right-click > Format Cells > Number > Number > change Decimal places   to 0.
Task 2: Change selected column   values   to be in font   Times New Roman,   Angle 45 degrees clockwise,   center horizontally and   vertically, and display 2 decimals
Task 3: Format   the   title, centered, bold, colored
Building charts
•            How can I create a chart?
o   Select data > Ribbon insert > select a chart.
The colored area shows   the plotted data; Move and resize the chart      Note: Click on   the chart, Extra Dynamic Ribbons appear: Chart Design and Format
•          How   to create a new   worksheet   with only   the chart?
o   Click on the chart > Ribbon Chart Design > Move Chart > New sheet.
•          How   to replace   x-axis labels   with   values in another column?
o   E.g., replace labels 1, 2,   …   , as   the students names
o   Click on the chart > Ribbon Chart Design > Select Data > Select Data Source, ( Legend   Entries panel shows   what info are contained in the plot, Horizontal (Category)   Axis
Labels shows   the   X-axis Label info) > Select Edit under Horizontal (Category)   Axis   Labels > select   the students name   values.
o   Warning! if   you accidentally select an additional cell, a   warning message   will pop-up   for   the above option!   – Check   the selected area
•          How   to   customize   the   title?
o   Click and change
•          How   to change   the color of   the bars?
o   Double-click one of   the bars> In Format Data Series > Fill and Line > Series
o   Highlight only one bar: select the bar and change its color
•          How   to change   the scale in   they-axis?
o   Double-click 代 写Advanced Excel: Excel Fundamentals review Module 1 – Excel interface and basic calculationsMatlab
代做程序编程语言  the   virtual label >   change Bounds and units
•          Transfer   from a   vertical bar chart   to a horizontal bar   chart.
o   Click on the chart > Ribbon Chart Design > Change Chart   Type > Bar >   vertical bar   chart.
•          How   to add data labels   for each bar?
o   Ribbon Chart Design >   Add Chart Element   > Data Labels
o   Double click   the Data Labels,   you can customize it on   the Format Data Labels
•          How   to change the chart background?
o   Double click   the background, select a picture on   the right panel Format Chart   Area
Part 2 - Formulas with Basic Calculations   Creating   formulas   to perform. calculations
File:   Module1   – Part 2 Using   formulas.xlsx
Worksheet:   Typing a   formula
•          Insert a   formula:   With =, e.g., = 3+4
•          Show   formulas: Ribbon Formulas > Formula   Auditing > Show   formulas
Worksheet: Doing math   with   cells
•         The basic operations: +, -, *,   /, ^,   Note:
o   #DIV/0   means that the formula is trying   to divide   something by 0.
o   3.5E+20 means 3.5 * 10^20
Worksheet: Comparing cells
Boolean   values: TRUE or FALSE   Note:
o   true, True, tRue,   … all recognized as   TRUE
o   Do not insert a space between < and =   for <=
o   Change   V1 and   V2   values,   the calculations are updated automatically.   Switch   to manually update   formulas:
o   Ribbon Formula > Calculation >Calculation Options, select   Manual.
o   Click Run calculation   to update   the   values.
Worksheet: Reusing   formulas
Task 1: Calculate   Average
o   L2: =(D2+E2+F2+G2)/4 r =AVERAGE(D2:G2)
o   Method 1: click I2, copy then paste in I3
o   Method 2: put   the mouse at   the bottom right corner,   when   the mouse becomes copy   handle (also called   fill down), propagate   to   the end/double click.
Task 2:   Adjust   for absence (each absence costs 2% of the avg)
o   J2: =I2-H2*2%*I2
Task 3: Check if the adjusted average is below or above the average of the class
o   Calculate the class average at J27: =AVERAGE(J2:J25)
Note:   To autofill   the entire   table requires absolute reference of cell   J27:
o   K2: =J2<$J$27
o   L2 :=J2>=$J$27   Note:
o   Absolute reference: Lock row or column, put $ correspondingly
o   Shortcut: F4
Task   4: Check if the student succeeds   with   the result of   the column "Above or equal   to average"
o   M2 : "=L2"
Naming range and cells
Worksheet: Naming range and cells
Ribbon Formulas, Function Library panel contains   tons of built-in   functions   Task 1: Calculate   the Min/Max/Average
o   K2: =MIN(G2:G25)
o   K3: =MAX(G2:G25)
o   K4: =AVERAGE(G2:G25)
Task 2: Name region G2:G25 as FinalMark:
o   Select region G2:G25, and rename by typing FinalMark in the Name Box
o   Re-do   the average calculation using the named region   K4: =AVERAGE(FinalMark)
Note: the FinalMark is an absolute reference, meaning $$.
When   you autofill move down   the   Average, the called region stays   the same as FinalMark   –   different if   you use =AVERAGE(G2:G25).
Task 3: Rename the region D2:D25 as Mid1Marks, E2:E25 as Mid2Marks,   F2:F25 as LabsMarks
Task   4: Check existing named regions:
Ribbon Fomulas > Defined Names > Name Manager.
Worksheet: Using cells   from dif. Sheets
B2: ='Doing math   with cells'!B2*'Comparing cells'!B2
Part 3   – Formulas with Basic and   Text Functions   Basic predefined   functions
Files:Module1   – Part 3 Basic and   Text Functions.xlsx

Function: = Function(Parameter 1, Parameter 2, …   , Parameter n)      result
Note: Not recommend to   type =SUM(B12+B13+B14…+B18), still   works but it means the   first   Parameter is a sum of all cells.
When   you add another row in between 2 and 3, =SUM(B12+B13+B14…+B18)   will not use the   added new row.
The SUM here is necessary. Use    =SUM(B12:B18).
When   you add another row in between 2 and 3, =SUM(B12:B18)   will use the added new row.
A   function   will normally return one single output, but some of   them may return a   table of   information.
=ROUND(B12:B18, 0)   will return all   values rounded   to   the nearest integer

Task 1:   The sum of   Table 1:
o   Method 1: Use   the SUM function   to add all   the cells from a range, separate by comma ,.   =SUM(C3,C4,C5,C6,C7,C8,C9,C10)
o   Method 2: Select the region using   the mouse/keyboard (Shift). =SUM(C3:C10)   Note:
The difference between Method 1 and 2: if   you add a new entity in row 5, the result by Method 1   will not include   the new entity.
Task 2:   The sum of   two regions, C3:C10 and F3:F10
o   Type =SUM(    Select the   first region      press Ctrl   while selecting the other region   Press   Enter   to see the result
Note: if   the   two regions overlap,   the   values in   the intersection region   will be used   twice.

•          Average C3,F3: =AVERAGE(C3,F3)
•          Average C3:C9:=AVERAGE(C3:C9)
•          Average C3:C9,F3:F9: =AVERAGE(C3:C9,F3:F9)   Note: Use Ctrl to select an additional region

•          Min C3:C9: =MIN(C3:C9)

•          Max C3   :C9   : =MAX(C3:C9)

Task: Calculate   the length of all names using   function LEN()
LEN(text) reports how many characters are in this text
o   B3: =LEN(B3) in B3, then autofill
Note:   The space   will be counted! E.g., B5 Pe   ter has 6 characters.

Task: Search   within a text   whether another   text is present
FIND(find   text,   within text, [start_num]) check if a subtext is   within a specific   text
o   [start_num]: at   which position I am going   to   start search
o   [] means optional parameter

Task: Extract 3 characters   from a text, starting   from position 4
=MID(Text from   where   to extract, beginning position, nr of characters)

Trimming: remove spaces at   the beginning and at   the end, as   well as   the duplicated spaces in   the middle
Task:   Trim   text in C2 and check the   trimmed and untrimmed   text length
Worksheets: Left and Right
=LEFT(Text, number of characters):
extract a certain number of characters starting   from   the beginning   =RIGHT(Text, number of characters):
extract a certain number of characters starting   from   the end

Task: Concatenate text 1   (C2) and   text 2 (C4)
o   Method 1 : =C2", "C4
o   Method 2 =CONCATENATE() : =CONCATENATE(C2,", ", C4)   Task: concatenate   text in a range
o   =CONCAT(C11:C15)
Task: concatenate   text in a range   with common separator symbol and ignoring blank cell
o   =TEXTJOIN(", ";TRUE;C11:C15)
Worksheets: Upper, Lower and Proper
Change   text into upper, lower and proper case   =UPPER(text), =LOWER(text), =PROPER(text)

Task: Get   the integer part of a   floating number
o   =Int(number):   to get the integer part of a floating point number
o   Int(2.5)    2; Int(-2.5)      -3
The resulting integer is always lower   than the query   value   Task: Round a   floating number
=Round(number, number of decimal places):   will round   the number to a specific number of   decimal places
Task:   Truncating a number   to a specific decimal places
=TRUNC(number, number of decimal places): truncates a number   to a specific number of   decimal places
Note: it is different   from rounding, as it only cuts   the   floating number up   to   the desired decimal   place

Task 1: Count the number of marks in D3:D26
=COUNT(Range1, Range2,…): to count   the number of   values   within one/many ranges.
Task 2: count the number of missing   values in D3:D26
=COUNTBLANK(Range):   will count all the empty cells   within ONE given range.
Task 3: count the number of last names in   A3:A26
=COUNTA(): count all the cells having a   value   in   whatever   types.
Note: =COUNT() only   work   with numerical numbers; =COUNT(A3:A26) gives   you 0

Task 1: Get   the initials of   the students in Column C
o   =LEFT(A2,1)LEFT(B2,1)
Task 2: Get   the   full names of   the students (LastName FirstName)   in Column D
o   =A2" "B2
Task 3: Calculate the average marks in Mid1: Final Exam
o   J2: =AVERAGE(F2:I2)
Note:   warning message Formula Omits   Adjacent Cells, because StudentID is also considered as   a number in Excel.
Task   4: In column K, display   the integer part of the average in column   J
o   =INT(J2)
Task 5: Display the rounded   value of   the average in column J, using   the number of decimal places   in O2
o   =ROUND(K4,$O$2)
Taks 6: Count the number of students in Column D (Full Name)
o   =COUNTA(D2:D25)
Task 7: calculate corresponding Minimum   values in columns F:J (Mid1:Average)
o   =MIN(F2:F25)
Task 8: Calculate Max,   Average, Sum, missing marks and number of marks in   each column
o   =AVERAGE(F2:F25)
o   =SUM(F2:F25)
o   =COUNTBLANK(F2:F25)
o   =COUNT(F2:F25)
Task 9:   Display   the results in J27:J30 in 2 decimals, and Center   the content in   F27:J32
o   ribbon Home > panel Number   to control the displayed decimals.
o   ribbon Home > panel   Alignment   to control   the display.
Part 4   – Formulas with Date and   Time   functions   Dealing   with dates
File: Module1   – Part 4 Formulas   with Date and   Time   functions.xlsx
Worksheet: Using Date   functions
To enter a date in a cell:
•         American: MM/DD/YYYY
•          French:   DD/MM/YYYY
Note:
o   Excel dates start on 01/01/1900,   to the end 12/31/9999.
o   If Excel recognizes   it is a date, it   will automatically align to   the right.
o   Try: 1/24/2000, 24/01/2000
o   when entering a   fraction, 1/3, excel   will regard it as a date 3-Jan.   You need =1/3   for   the   value.
In   the system,   the date   will be stored as a serial number.   Try to see a date in number   format.
Task 1: Todays date
o   =TODAY()
Task 2:   Todays date and   time
o   =NOW()
Task 3: Return the Day, Month,   Year, and day of the   week of Cell B3
o   =DAY(B3), =MONTH(B3), =YEAR(B3), =WEEKDAY(B3)
Note: =WEEKDAY(Date, [option]) default   Sunday being number 1.
Task   4: Create a date from   values of Day (13), Month (3), and   Year (2018)
o   =DATE(DD,MM,YYYY)
Task 5: how many days   between 01/01/2020 and 01/02/2020 (including start and end dates)
o   =B14   – B13 +1
Task 6: Calculate   the date   that   X days after a certain date.
o   Adding one day to   the   date, =B13+1
Task 7: calculate   the number of   years between two dates.
o   =DATEDIF(B17,B18,"Y")
Dealing   with   times

To insert a   time (in 24h format): HH:MM:SS   or HH:MM
Like   the date, Excel   will convert it into a series of numbers.   You can check   the associated   value in   Number   format, e.g., 12:08    0.51.
   
Task: Calculate duration between   two   times (B5).
o   =B3   – B2
Task: Display   time now in a different   format.   Task: Extract different part of a time
o   E2: =NOW(), then change   format   to   Time
o   E3: =HOUR(E2)
o   E4: =MINUTE(E2)
o   E5: =SECOND(E2)
o   Click Calculate Now   to update
Task: Build   time   from Hour, Minutes and Secondsin G2:G5
o   =TIME(H2,H3,H4)
Task: Calculate   the duration between   the corresponding start and end   time in columns B and C.
o   In D9, =C9   – B9 then Change the format   to 13.30   to remove   AM and Propagate by   Tips to propagate:
o   Double click   the copy handle   OR
o   Select   the D9:D24,   then press shortcut Ctrl + D   OR
o   Select   the D9:D24, select ribbon Home Panel Editing Down

Task: calculate   total duration in D9:D24
o   =SUM(D9:D24)
Note: If   you see 15:10, it is   wrong! Because it exceeds the 24hr.   A   wrong display   format! Change   to 37:30:55! Date   format
Task: calculate   total payment   when the pay per hour is 30$
   
   
Wrong   way: only multiply   the   total duration and   the hourly pay
Excel understands   the duration as 1.62 days (check it in Number   format).   We need   to multiply it by 24 hr/day.
o   =D25*D26*24

         
加QQ：99515681  WX：codinghelp  Email: 99515681@qq.com
