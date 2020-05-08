# 1.Star Counter

## Instructions

* Create a VBA Script that tallies the number of "Full Stars" per row and enters them into the Total column. Starter Code is provided, but feel free to start from scratch if you want an extra challenge :-)

* **Bonus**:

  * Instead of hard-coding the last number of the loop, use VBA to determine the last row automatically (i.e. do not use for i = 2 to 51)

  * Create two charts: 

    * One to see if there is a relationship between Program Type and Rating (Bar Chart)

    * The other to see if there is a relationship between Date and Rating (Line Graph)

* **Hint**:

  * You will need to use a nested for loop.

  * You will need to create a variable to hold the number of stars and continually reset this variable at the start of each row.

# 2.Gradebook

## Instructions

* Using `grader.xlsm` as a starting point, create a grade calculator using **conditionals**. This calculator will convert a student's numeric grade into a letter grade, and style the resulting cell accordingly.

* Once complete your script should perform the following:

  * If the score is over 90, the student will receive an "A" in the letter grade cell, and the Pass/Warning/Fail cell will be filled green with the text "Pass."

  * If the score is between 80 and 89 (inclusive), the student will receive a "B" in the letter grade cell, and the Pass/Warning/Fail cell will be filled green with the text "Pass."

  * If the score is between 70 and 79 (inclusive), the student will receive a "C" in the letter grade cell, and the Pass/Warning/Fail cell will be filled yellow with the text "Warning."

  * Finally, if the score is below a 70, the student will receive an "F" in the letter grade cell, and the Pass/Warning/Fail cell will be filled red with the text "Fail."

## 3.BONUS

* Create a second button that resets the grades to the original state and then establishes the previous grade in a row labeled "Last Grade."

# Checkerboard layout

## Instructions

* Using VBA scripts, create an 8x8 grid with alternating red and black squares.

## Hints

* You will need to use nested for loops, conditionals, mods, and formatting to create the board.

* This is a tricky problem! Try to pseudocode a plan first.

  * Unlike previous activities, this activity can be solved in a multitude of different ways. While some methods may be more efficient than others, simply finding a solution to the problem is a great start!

# 4.Credit Card Checker

## Instructions

* Create a VBA script that will process the credit card purchases, identifying each of the unique brands listed.

* For the _Basic_ assignment, create a single pop-up message for each of the Credit Card brands listed by looping through the list.

* For the _Advanced_ assignment, tally the total credit card purchases for each Credit Card brand and add it to the summary table.

## Notes

* This assignment is extremely similar to the basic version of the homework assignment. So let's buckle down and analyze this!

# 5.Wells Fargo Part I

## Instructions

1. Extract words before the phrase "\_Wells_Fargo" to figure out the state.

2. Add the state to the first column of each spreadsheet.

3. Convert the headers of each row to simply say the year.

4. Convert the numbers to currency values for all cells.

* **Hints**

  * First work on getting the correct formatting on one sheet before moving onto creating a loop that formats each sheet within your workbook.

  * If you are looking for a useful resource for finding the code to loop through all worksheets in a workbook, check out this link [here](https://support.microsoft.com/en-us/help/142126/macro-to-loop-through-all-worksheets-in-a-workbook)

Data Source: [Wells Fargo Bank Deposit Data](https://www.datazar.com/project/p54404488-d82b-49f5-b43e-d63447daee32/files)

# 6.Wells Fargo - Part II

* In this second part of the mini project project, you will be combining all of your previous Excel sheets into one massive table on a new sheet.

**Instructions**

* Loop through every worksheet and select the state contents.

* Copy the state contents and paste it into the Combined_Data tab