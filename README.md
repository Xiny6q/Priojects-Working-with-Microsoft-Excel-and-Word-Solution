Download link :https://programming.engineering/product/priojects-working-with-microsoft-excel-and-word-solution/

# Priojects-Working-with-Microsoft-Excel-and-Word-Solution
Priojects Working with Microsoft Excel and Word Solution
This description of the assignment contains instructions that tell you to create files with names that have a specific format. In these file names; the “youraccountname” is your UWO username.

Project 1: Microsoft Excel

This project prepares a monthly revenue information ( profit/loss ) for salespersons in your company for the month of April, 2019. The partially completed workbook used for this project is stored in the file “Commissions.xlsx”. You must use the supplied file “Commissions.xlsx” as your workbook or you will lose major marks if you use any other file.

The Net Earnings for each Salesperson minus the Per Unit Costs to the Salesperson will compute the profit or loss for each individual Salesperson.

Your finished work should resemble the image below exactly (but with the correct answers of course… and the colors and column spacings do NOT have match exactly, close enough is good enough in regards to the actual colors and widths. So, do not get too worried about it…)


Each salesperson sells two of the items that your company manufactures. This is the ‘Niblick’ and the ‘Pit Mashie’. Each salesperson also offers the customer an optional Warrantee on each item they sell. This Warrantee does not cost the Salesperson any money, so the commission on the warrantees is pure profit for the salesperson. This incentivizes them to sell more Warrantees. The Warrantees are the same cost for either product.

The workbook is intended to compute the revenue profit/loss for each salesperson.

The Product Sales Revenue includes ONLY the total from the sales of the two products (the ‘Niblick’ and the ‘Pit Mashie’). It is based on the number of units sold times the price to the consumer for each product.

The Product Sales Commission the salesperson earns is based on the total sales of ONLY both of the items. This column does not include the commissions from the Warrantee sales. This commission is based on the sliding scale shown in the table starting at cell B21.

The Warrantee Sales Revenue is based on the total number of warrantees sold times the warrantee cost to the customer.

The Warrantee Sales Commission the salesperson earns is based on the total sales of ONLY the number of warrantees sold that month. This column does not include the commissions from the Product sales. This commission is based on the sliding scale shown in the table starting at cell F21.

The salesperson’s Net Earnings for the month is the total of all the commissions they earned in the month.

The Unit Cost to the Salesperson is the total cost the salesperson paid your company for the two items they sold in the month based on the table starting at cell F15.

The numbers in the tables may change month by month, so the current values are stored in the worksheet, BUT you must use cell referencing in ALL your equations. (hint: you cannot use the actual numbers in the calculations)

The Revenue Profit/Loss is then simply computed as the difference between the Net Earnings and the Unit Costs to the Salesperson.

All calculations must be written so they can be copied to each corresponding cell using the Absolute and Relative cell referencing as covered in class. (i.e. write the formula in cell E4 so as it can be copied to cells E5 to E11 without any changes)

Complete the following instructions and save your workbook in a file named:

“youraccountname_Commissions.xlsx” and attach this file to your submission.

Develop the formulas for the Product Sales Revenue, Product Sales Commissions, Warrantee Sales Revenue, Warrantee Sales Commissions, Net Earning, Unit Costs to the Salesperson and the Revenue Profit and/Lost columns. The formulas for these calculations will use the information provided in the Purchase Price to Customer, Product Sales Commission Scale, Per Unit Costs to Salesperson and Warrantee Commission Scale tables of the worksheet. The formulas must be created using cell references.

Copy the formulas for the first salesperson (Isabelle Ringing) to the remaining seven salespersons.

Calculate the Totals row (row 12) using the values in each associated column.

Delete any unused sheets in the workbook.

Format the worksheet as follows:

Display all dollar amounts with the Canadian currency symbol and two decimal places

Display all percentages as percent (%) values as in the image above.

Change the name of the worksheet to April 2019

Change the first row so it is spread out over all the used columns (A through K) and center the name of the company in row 1. Bold and Increase the font size of the Company Name Title to 14 or 16 (your choice).

Highlight (change the background color) and bold the titles of the row of Labels (row 3), the row of Totals (row 12)

Make sure to add the grid lines so they appear for the row of Labels (row 3) and the row of Totals (row 12) and the Profit/Lost column (Column K) exactly as in the image above.

Highlight (change the background color) and bold the titles of the table labels (rows 15 and 21) exactly as shown in the image above. (note – the highlighting is JUST in the table titles in these two specific rows as in the image above.

Put a bold line around each of the table as shown in the image and make sure the grids are set as in the image above.

Adjust the column sizes to fit the information contained in them. All values must be displayed in all resolutions (no ####### due to too narrow cell spacing).

Add your name to the worksheet following “Prepared by:” (B28)

Add the current date (format Month Day, year – example: July 6, 2023) to the worksheet following “Date:” (K28)

Change the Revenue Profit/Loss column so the text is black and regular and the background is light yellow for positive values and text is yellow and italics and the background is red for negative values. You must set this so the colors will automatically change if the values change from positive to negative or from negative to positive. (note: image above does not have the correct answers…

notice some of the values are different from the actual assignment file)

Project 2: Microsoft Excel

payments between the amount that will go towards paying off the actual principle and amount of the payment that goes towards the interest payment. Notice that your early payments are going almost entirely to paying the interest of the loan. Conversely, the later payments are counted more towards paying the principal of the loan. The banks make sure that they make their money up front. You must use the supplied file “MortageCalculator_BLANK.xlsx” as your workbook or you will lose major marks if you use any other file.

The partially completed workbook contains two worksheets to provide the information necessary to complete this project.

Complete the following instructions and save your workbook in a file named “youraccountname_MortageCalculator.xlsx” and attach this file to your submission.


Use cell references not cell values in all of the formulas.

On the first worksheet:

Calculate the amount borrowed in Cell B6.

Develop the formula to calculate interest rate of the loan in cell F4. The formula must use the VLOOKUP function. It will base the Interest Rate on the value entered in cell F5 (the duration of the loan in years) based on the table in the Information worksheet of this workbook.

In Cell F7, calculate the total number of loan payments (# Payment Periods).

In Cell B11, calculate the monthly payments based on paying at the beginning of the period Monthly payment

calculated using the PMT function

using cell references not cell values

shown as a positive number

payment at beginning of payment period

In Cell B12, calculate the monthly payments based on paying at the end of the payment period Monthly payment

calculated using the PMT function

using cell references not cell values

shown as a positive number

payment at end of payment period

In Cells C11 and C12 calculate the total amount of the loan based on the corresponding payment structure. (C11 total amount paid based on beginning of period payments – C12 based on end of period payments.)

Loan payments are structured on a sliding scale of principle (what you borrowed) and interest (what you pay in order to borrow the money). The scale starts with more of the payment going towards interest than towards the principle borrowed.

The Cells B17 and B18 show the breakdown of how much of the payment is principle and how much is interest in the first payment.

The following cells will show the breakdown at the quarter, half, three quarters and last payment.

Cell C16 will calculate the payment number at the quarter point (25% point of paying back the loan). This is simply the total number of payments multiplied by 25%. D16 will then be the payment number at the half way point (50% point of paying back the loan). This is simply the total number of payments multiplied by 50%. E16 will then be the payment number at the three quarter point (75% point of paying back the loan). This is simply the total number of payments multiplied by 75%. Finally, F16 is the last payment. This will just be the total number of payments represents the value of the number of the last payment made.

Cells B17 and B18 have been prefilled with the formulas for the breakdown of principle and interest. Edit these formulas so they can be copied to the corresponding cells (C17 to F18)

In cells B20 to F20 simply compute the totals of the principle and interest for each column to demonstrate that the combination of the two do indeed add up to the monthly payment (based on paying at the end of the period)

Format the first worksheet as follows;

Display all dollar amounts with currency symbol and appropriate decimal places

Merge and Center the label in A1 up to C1. Increase the font size and bold the label

Put a box outline in the range of A1 to C13 and underline the label in A1

Merge and Center the label in E3 up to F3. Increase the font size and bold the label

Put a box outline in the range of E3 to F7 and line all the boxes in that range

Add your name to the prepared by in Cell B23

Highlight (change the background color) the cells in the range of A15 up to F15 and the range of A20 to F20 as shown in the image above.

Put a box outline in the range of A15 to F18 and line all the boxes in that range

Highlight (change the background colors ) of all the four (4) label rows so they look like the image above. (i.e. the dark red, red, blue and light orange cells).

Change the font color of any cell that can be changed by the user to a blue color text.

Cells B4, B5, F5 and F6

Rename the worksheet labeled “Sheet1” to “Payments”

remove any extra (unused) worksheets from the workbook

Project 3: Microsoft Word and Excel

For this project, create a Word and an Excel document that contain information about a house (Residence or Condominium) you are planning to purchase. The document should be used to provide information to the financial institution you are approaching to provide the funds needed to purchase this item. The information contained in the Word document and the Excel spreadsheet can be real or fictional. This should be three or four short paragraphs in length.

You are not graded on what you write, only that the imbedded Excel is included with some text.

Describe the location of the house, the layout, how many rooms and bathrooms, etc.

Complete the following instructions. Save your Word document in a file named:

“youraccountname_MortageCalculator.doc(x)”.

Link the mortgage portion of Project 2 (Cells A1 to C13) to that document.


MORTGAGE CALCULATION TABLE

House Price

$1,245,450.00

Down Payment

$320,000.00

Amount Borrowed

$925,450.00

Total

Monthly

Amount

Payments

Paid

Beginning of Pay Period

$6,587.49

$2,371,497.44

End of Pay Period

$6,630.04

$2,386,813.37

In the Excel workbook, select all of the cells in your spreadsheet containing data and copy the selected range to the clipboard. Open the Word document you created and insert your Excel workbook into it by using the paste option that allows the Excel workbook to be linked into the Word document.

NOTE: This is a ‘live’ link. So data changed in the Excel sheet will instantly be updated in an opened Word document or will be reflected next time the Word document is opened. These updates will occur with no intervention, editing or change is required by the user. The TA will test this link by making a change in your Excel file and then checking your Word file. The values in the Word file MUST reflect the change in the Excel file.

You can test your link by changing values in the Excel document and then checking your Word document.

Project 4: Information Systems Questions about Your Company

Create a one page MS Word document and complete the following questions pertaining to the business you described in Assignment One (1).

1.) Would allow a committee of your employees to make decisions or should you as the company owner have the final vote and make all the decisions?

– briefly explain your answer

2.) Name one way you might use MS Excel in your company?

– briefly describe how it could be used and what need it would fulfill.

3.) After what you have learned in CS1032, do you think you would be interested in running a company of your own?

– briefly explain your answer

The format of this document should be identical to format you used in Assignment One (1).

Place your name, followed by the company name at the top.

Fill in the required information after.

At the end of the document, include your name, Student number and Western ID (the first part of your

Western email (i.e. if your email is – dernt373@uwo.ca your ID will be – dernt373)

Formatting is not important as long as the document is easy to follow:

This document must be a Word file saved and submitted as a .doc (or .docx) file

The name must be a combination of your Western Account Name and the name of your company. The file name must be youraccountname_companyname_A6.doc (or .docx)

– example (from above) dernt373_MaggicSoftware_A6.docx

Submission Instructions:

Upload and submit the following 4 files using the assignment tool on the CS1032 OWL site:

youraccountname_Commisssions.xlsx (for later versions) (.xls for earlier version)

youraccountname_MortageCalculator.xlsx (for later versions) (.xls for earlier version) youraccountname_MortageCalculator.docx (for later versions) (.doc for earlier version) youraccountname_companyname_A6.docx (.doc for earlier version)

AND FOR THE FINAL TIME, PLEASE: Do not cheat or copy or work in groups. Remember, you will need to know how to perform these functions for the final exam. Remember, you must achieve at least a 49% on the final exam in order to receive a passing grade
