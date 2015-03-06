---
layout: default
title: "PRDV252: Intermediate Excel"
course_description: "This class is intended for students who have a basic understanding of spreadsheets and are now ready to delve deeper into formatting, formulas and functions, multi-page spreadsheets, charting data, creating tables that have database features, and be introduced to pivot tables."
next: ../Unit10
previous: ../Unit08
---
**Unit 9: Multi-Page Spreadsheets** <span id="9"></span> 
*In this unit, you will learn how to work with large workbooks, and this
unit will highlight various features that will make working in Excel®
more productive. We will explore formulas and functions that drill
downthrough multiple worksheets for their answers. When a workbook gets
large, it is important to understand how to format each worksheet tab so
that information is easier to find. You learned about formatting
worksheet tabs in Subunit 2.2.2. If needed, reread that section. *

**Unit 9 Time Advisory**  
Completing this unit should take you approximately 3 hours.  
  
 ☐    Subunit 9.1: 0.75 hours  
  
 ☐    Subunit 9.2: 0.75 hours  
  
 ☐    Subunit 9.3: 1.5 hours

**Unit9 Learning Outcomes**  
Upon successful completion of this unit, you will be able to:
-   describe how to write a formula that drills down through multiple
    worksheets;
-   create and format multiple worksheets at one time; and
-   describe how to group and ungroup worksheets.

**9.1 3D Formulas** <span id="9.1"></span> 
*3D formulas are used to combine information from multiple worksheets
onto one worksheet for an answer. Information can be taken from as many
worksheets as needed. When using a 3D formula, it is easier if your
worksheets are adjacent to each other in the workbook. It makes it
easier because the syntax of the formula is easier to read. Each
worksheet needs to be named in the formulas if working with non-adjacent
workbooks, and this can get cumbersome to write as well as read. Let’s
say you are working in Sheet 7 and want to add two numbers from two
other worksheets in your workbook. The 3D formula will look like this:
=Sheet1!C3+Sheet4!$C$6. This will add the number in Sheet 1 cell C3 to
the number in Sheet 4 cell C6 and place it in the worksheet you are in,
Sheet 7. The exclamation point (!) is needed in the formula at the end
of each sheet name followed by the cell reference, which can either be
relative or absolute. You learned about relative (C6) and absolute
($C$6) referencing in Unit 3. If you want to add all of the cells B3
from all worksheets in the workbook and they are adjacent to each other,
it would look like: =sum(Sheet1:Sheet6!$B$3). This formula tells us that
it is adding each number in B3 in each of the worksheets 1 through 6.*

-   **Web Media: The Saylor Foundation’s “3D Formulas”**
    Link: The Saylor Foundation’s [“3D
    Formulas”](http://www.youtube.com/watch?v=1gDg7MVCeZ8) (YouTube)  
        
     Instructions: Watch the video about 3D formulas. When you are
    finished, open up your own spreadsheet software and try to create
    some 3D formulas.  
        
     Completing this assignment should take approximately 45 minutes.

**9.2 Formulas Across Workbooks** <span id="9.2"></span> 
*Now that you’ve created 3D formulas within one workbook using various
worksheets, let’s look at creating a formula that uses two workbooks.
This is a little bit more complicated because you have to be cognizant
of where the workbooks are saved and if they are linked. If the
workbooks are linked, they shouldn’t be moved once the formula is
created or the formula may not work. If they are moved, both need to be
moved so that the link can stay useful. When you create formulas between
workbooks, you can either link them or just create the formula without
linking them. But, if they are* not*linked and you change or update a
number in one workbook, the formula you created will not be updated with
the new information. This may make more sense after you have watched the
video.*

-   **Web Media: The Saylor Foundation’s “Formulas Across Workbooks”**
    Link: The Saylor Foundation’s [“Formulas Across
    Workbooks”](http://www.youtube.com/watch?v=WH6YWZ1o6Y0) (YouTube)  
        
     Instructions: Watch the video about 3D formulas across workbooks.
    When you are finished, open up your own spreadsheet software and try
    to create a linked formula between two workbooks. You can use
    spreadsheets that you created earlier in this class.  
        
     Completing this assignment should take approximately 45 minutes.

**9.3 Unit 9 Exercises** <span id="9.3"></span> 
-   **Activity: The Saylor Foundation’s “Practice Exercise for Unit 9”**
    Link: The Saylor Foundation’s [“Practice Exercise for Unit
    9”](http://www.saylor.org/site/wp-content/uploads/2013/10/PRDV252-Unit-9.3-Instructions-FINAL.pdf) (PDF)  
        
     Instructions: Follow the instructions on the PDF. You will need the
    accompanying Excel® file below to complete this exercise.


