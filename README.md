# Payment-History-Portal
This spreadsheet acts as a password protected payment history portal
## Situation
* Customers at a learning academy did not have access to their payment history for their individual student. As a result, they were often confused about past payments and were not sure when to expect when another Payment was due.  Since giving them access to their personal history would enable them to have knowledge about their past payments, a new portal would need to be created for access.
## Task
**A Few Things to Consider**
* Since payment information is meant to be kept private, each customer would need a passkey as identification and to allow them into the spreadsheet
* The information needed for this spreadsheet is already recorded and stays up to date in _another_ spreadsheet. (For this GitHub example, I've chosen to put the data within the same spreadsheet) Then the information is then queried to the current sheet.
## Method
* Have an accurate and up to date record of all student lessons and payments
* Use Google Apps Script to:
  * Generate an alert that prompts the user for a passkey
  * Create a new blank spreadsheet where specific information based on the passkey entered will populate the rows
  * Enable each user to see their payment history in a efficient and cohesive manner
### Google Apps Script
[See the Google Apps Script functions]()

**Functions created in PasswordProtect**
*idPopUp()
 - Generates an alert that prompts the user to insert a Student ID. The alert will remain until the user provides a valid Student ID
 -  The valid IDs are stored within this function
 -  When a valid ID is entered, idPopup calls the function insertSheet(), inserts the valid ID in a determined cell, sets the name of the new sheet to the student ID, and imports an image that functions as the finishSession button
* insertSheet()
 - Creates the new sheet and sets formulas in each cell that are responsible for quering/importing the data from the other sheet
    - In the original example, the inforrmation is queried from a sheet OUTSIDE of the current spreadsheet
 - 

## Sample Sheet and Results
[USE THE SAMPLE SPREADSHEET HERE](https://docs.google.com/spreadsheets/d/1vE5K8VghNGc4zUb-PTDjudg1maEpeeZihRpkhA2uFEc/edit?usp=sharing)

The Home Page (Home Page - Start Here) gives the user instructions on how to navigate the spreadsheet. In order to start the script that runs the **START CALCULATOR** button, the user must use a signed in Google Account and accept the permissions of the new application.

After doing so, the user can copy a Student ID from Sheet 2 (Student_IDs), generate the alert box with the button, then paste the copied ID into the prompt. A new sheet will be created that holds the data pertaining to the student whose ID was chosen. 

When the user is finished with the new sheet, they must click the X button in cell A3 to close the session and delete the page. They will be returned to the home page. 
