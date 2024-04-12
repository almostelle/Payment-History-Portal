# Payment-History-Portal
This spreadsheet uses Google Apps Script to function as a password protected payment history portal
## Situation
* Customers at a learning academy did not have access to their payment history for their individual student. As a result, they were often confused about past payments and were not sure when to expect when another payment was due. In order to give them access to their personal history, a new portal would need to be created that ensured both accuracy and privacy.
## Task
**A Few Things to Consider**
* Since payment information is meant to be private, each customer would need a passkey to allow them to access only their information
* The records needed for this new spreadsheet (I'll call it the *user* sheet) are already recorded and stay up to date in _another_ spreadsheet (I'll call it the *records* sheet).* Then, the information is queried to the user sheet.<br>

  \*For this GitHub example, I've chosen to put the records sheet, called Sample_Data, in the same spreadsheet as the user sheet. In the original example, the records sheet is located in a different spreadsheet.
## Method
* Have an accurate and up to date record of all student lessons and payments
* Use Google Apps Script to:
  * Generate an alert that prompts the user for a passkey
  * Create a new user spreadsheet where specific information based on the valid passkey will populate the rows
  * Enable each user to see their payment history in an efficient manner
### Google Apps Script
[See the Google Apps Script functions](https://github.com/almostelle/Payment-History-Portal/blob/main/PasswordProtect.js)

**Functions created in PasswordProtect**
* **idPopUp()**
  - Generates an alert that prompts the user to insert a Student ID. The alert will remain until the user provides a valid Student ID
  -  Stores an array of the valid IDs
  -  Calls the function insertSheet() when a valid ID is entered, from which a new sheet is created
  -  Inserts the valid ID in A2
  -  Sets the name of the new user sheet to the student ID
  -  Imports an image that functions as the finishSession button then assigns the appropriate script
 
* **insertSheet()**
  - Creates the new user sheet and sets different formulas in each cell that are responsible for quering/importing the data from the records sheet
    - In the original example, the information is queried from a records sheet OUTSIDE of the current spreadsheet
  - Calls the function formatSheet()
 
* **formatSheet()**
  - Stores an array of strings for column titles
  - Sets the values of a range to the array titles
  - Formats the specified cells' background and text colors
 
* **finishSession()**
  - Generates an alert that prompts the user to answer a question ("Do you want to close the session?")
  - If the user selects "Yes", the user sheet is deleted

## Sample Sheet
[USE THE SAMPLE SPREADSHEET HERE](https://docs.google.com/spreadsheets/d/1vE5K8VghNGc4zUb-PTDjudg1maEpeeZihRpkhA2uFEc/edit?usp=sharing)

The Home Page (Home Page - Start Here) gives the user instructions on how to navigate the spreadsheet. In order to start the script that runs the **START CALCULATOR** button, the user must use a signed in Google Account and accept the permissions of the new application.

After doing so, the user can copy a Student ID from the sheet Student_IDs, start the script that generates the alert box with the button, then paste the copied ID into the prompt. A new user sheet will be created that holds the data pertaining to the student whose ID was chosen. 

When the user is finished with the new user sheet, they must click the X button in cell A3 to close the session and delete the page. They will be returned to the home page. 

## Results

