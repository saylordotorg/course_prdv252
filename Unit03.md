**Unit 3: Mathematical Computations** <span id="3"></span> 
*This is where the fun begins. Excel®’s usefulness shines in the way it
handles formulas and functions. Formulas are mathematical equations that
you write yourself. Functions are the prebuilt formulas that Excel® has
available for use, such as finding the average or standard deviation for
a set of numbers. A new vocabulary word you need to watch for is RANGE.
A range is a set of numbers that are used in a formula or function. The
range is usually, but not always, contiguous numbers, which are next to
each other in the spreadsheet. Also keep in mind that most often cell
references are used, not actual numbers. For example, if you were adding
two numbers, say 12 and 15, and the first number is in cell A3 and the
second in cell B3, then the formulas would look like: =A3+B3. It would
not look like: =12+15. Cell references are used so that later if you
wanted to change the number in A3 all of the associated formulas and
functions that use that number would automatically update, even if they
are on a different worksheet or in a different workbook!  
    
 Each of the subunits in this unit builds on the next. Be advised that
you should not skip units as you may write a formula in Subunit 3.1 that
you will need to revise or add to in Subunit 3.3.  
    
 There is a lot of information in this unit so don’t rush through it. It
may be a bit confusing at first but keep in mind that this is the third
unit of this Excel® class but it uses the second chapter of the
textbook. So you may be working in classroom Subunit 3.2 but it is
textbook Section 2.2.*

**Unit 3 Time Advisory**  
Completing this unit should take you approximately 3 hours.  
  
 ☐    Subunit 3.1: 30 minutes  
  
 ☐    Subunit 3.2: 1 hour  
  
 ☐    Subunit 3.3: 1 hour and 10 minutes  
  
 ☐    Subunit 3.4: 20 minutes

**Unit3 Learning Outcomes**  
Upon successful completion of this unit, you will be able to:
-   create basic formulas;
-   explain cell referencing;
-   explain the order of mathematical operations that Excel® follows in
    a formula;
-   identify formula auditing tools;
-   use relative and absolute cell references in formulas; and
-   use the COUNT, AVERAGE, MAX, MIN, PMT, and FV functions.

**3.1 Formulas** <span id="3.1"></span> 
-   **Reading: *How to Use Microsoft Excel*®: “Section 2.1”**
    Link: *How to Use Microsoft Excel®*:
    [“](https://resources.saylor.org/wwwresources/archived/site/textbooks/How%20to%20Use%20Microsoft%20Excel.pdf)[Section
    2.1](https://resources.saylor.org/wwwresources/archived/site/textbooks/How%20to%20Use%20Microsoft%20Excel.pdf)[”](https://resources.saylor.org/wwwresources/archived/site/textbooks/How%20to%20Use%20Microsoft%20Excel.pdf)
    (PDF)  
        
     Instructions: Go to Chapter 2: Mathematical Computations and find
    Section 2.1: Formulas, which starts on page 104. Then, open the link
    in the next Resource Box for the follow-along Excel® file to work on
    as you read the chapter.  
        
     Completing these assignments should take approximately 30
    minutes.  
        
     Terms of Use: The textbook used in the link above is released under
    a [Creative Commons Attribution-NonCommercial-ShareAlike 3.0
    Unported
    License](http://creativecommons.org/licenses/by-nc-sa/3.0/) without
    attribution as requested by the work’s original creator or licensee.

-   **Activity: The Saylor Foundation’s “Excel® Text Chapter 2”**
    Link: The Saylor Foundation’s [“Excel® Text Chapter
    2”](https://resources.saylor.org/wwwresources/archived/site/wp-content/uploads/2013/10/Excel-Text-Chapter-2.xls) (XLS)  
        
     Instructions: This Excel® file is in the earlier version of Excel®
    with the file extension .xls. If you have the newer version of
    Excel®, open it up and resave it in the current version of Excel®
    which has the file extension of .xlsx. There are features that the
    newer version has that you may want to play with so you don’t want
    to leave it in Compatibility Modeview. If you have a newer version
    and leave it as an .xls, you will see the text [Compatibility Mode]
    on the Title Bar. This will not affect the document at all. But
    remember, the book will be showing the newer version in the pictures
    that accompany the text. As this chapter deals with formulas and
    functions, all Excel® versions write formulas and functions the
    same. This Excel® workbook has several worksheets. You will complete
    all of the worksheets while working through Unit 3. Save the file
    often so that you don’t lose any of your work.  
      
     Completing this activity should take approximately 1 hour.

**3.2 Statistical Functions** <span id="3.2"></span> 
-   **Reading: *How to Use Microsoft Excel*®: “Section 2.2”**
    Link: *How to Use Microsoft Excel®*:
    [“](https://resources.saylor.org/wwwresources/archived/site/textbooks/How%20to%20Use%20Microsoft%20Excel.pdf)[Section
    2.2](https://resources.saylor.org/wwwresources/archived/site/textbooks/How%20to%20Use%20Microsoft%20Excel.pdf)[”](https://resources.saylor.org/wwwresources/archived/site/textbooks/How%20to%20Use%20Microsoft%20Excel.pdf)
    (PDF)  
        
     Instructions: Read Chapter 2, Section 2.2: Statistical Functions
    that starts on page 120. You will continue to use the file you
    started above (Excel® Text Chapter 2). The functions in this chapter
    are the basic ones that you should master in Excel®. Once you are
    finished with the chapter, practice using these functions in a
    spreadsheet of your own making.  
        
     Completing this assignment should take approximately 1 hour.  
        
     Terms of Use: The textbook used in the link above is released under
    a [Creative Commons Attribution-NonCommercial-ShareAlike 3.0
    Unported
    License](http://creativecommons.org/licenses/by-nc-sa/3.0/) without
    attribution as requested by the work’s original creator or licensee.

**3.3 Functions for Personal Finance** <span id="3.3"></span> 
*In this section, you will finish Chapter 2 of the text. It introduces
functions for personal finance. You will learn the PMT and the FV
function. The PMT, or Payment function, is a useful function and can be
used with all personal loans. You can work out payment options using the
variables of interest rate and time duration of the loan so you can make
better decisions when acquiring a loan. The one item that you need to
remember is that loans are usually based on an annual interest rate.
Since we usually make payments on a monthly basis, the annual or yearly
rate must be divided by 12. So if the interest rate is 6%, that rate
should be 6%/12 in the payment function because you only pay a portion
of the 6% each month (which would be 0.005% a month). In addition when
you are deciding on the length of the loan, if it is a 10-year loan,
when talking about the payment, you would need to multiple 10 years
times 12 months since you pay monthly. Or you could just use 120 months
as the length (10\*12). Either way, the length, or Nper in the function
box, needs to reflect monthly payments. No one wants to make a car
payment once a year – ouch! These are often forgotten by someone just
learning the PMT function.*

-   **Reading: *How to Use Microsoft Excel*®: “Section 2.3”**
    Link: *How to Use Microsoft
    Excel®*: [“](https://resources.saylor.org/wwwresources/archived/site/textbooks/How%20to%20Use%20Microsoft%20Excel.pdf)[Section
    2.3](https://resources.saylor.org/wwwresources/archived/site/textbooks/How%20to%20Use%20Microsoft%20Excel.pdf)”
    (PDF)  
        
     Instructions: Read Section 2.3: Functions for Personal Finance,
    beginning on page 144. This subunit will finish Chapter 2 of the
    text. You will continue to use the file you used in Subunit 3.2
    above (Excel® Text Chapter 2).  
        
     Completing this assignment should take approximately 1 hour and 10
    minutes.  
        
     Terms of Use: The textbook used in the link above is released under
    a [Creative Commons Attribution-NonCommercial-ShareAlike 3.0
    Unported
    License](http://creativecommons.org/licenses/by-nc-sa/3.0/) without
    attribution as requested by the work’s original creator or licensee.

**3.4 Exercises for Unit 3** <span id="3.4"></span> 
-   **Web Media: GCFLearnFree.org’s “Excel® 2010: Creating Complex
    Formulas”**
    Link: GCFLearnFree.org’s
    [“](http://www.gcflearnfree.org/excel2010/9)[Excel® 2010: Creating
    Complex
    Formulas](http://www.gcflearnfree.org/excel2010/9)[”](http://www.gcflearnfree.org/excel2010/9)
    (HTML)  
        
     Instructions: Go through the 6 pages, watching the video on page 2.
    Pay particular attention to the explanation of Order of Operations.
    This is what Excel® uses to calculate formulas. When you have
    finished, there is a practice document at the end of this
    presentation. If you have a newer version of Excel®, open the
    spreadsheet and follow the instructions to practice creating
    formulas.  
        
     Completing this assignment should take approximately 20 minutes.  
        
     Terms of Use: Please respect the copyright and terms of use
    displayed on the webpage above.


