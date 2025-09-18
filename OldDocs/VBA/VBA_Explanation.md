# VBA Scripts Explanation for Interface.xls

This document provides an overview of the VBA scripts contained within the `Interface.xls` workbook, which serves as the main application for the Production Control System. The scripts are divided into modules and forms, each responsible for a specific part of the system's functionality.

## Modules (`.bas` files)

These files contain reusable functions, subroutines, and public variables that are accessed by the forms and other modules.

- **`a_Main.bas`**: The entry point of the application. The `ShowMenu` subroutine initializes and displays the main user interface (`Main.frm`).
- **`a_ListFiles.bas`**: Contains the `List_Files` function, which populates list boxes on the forms with files from specified directories (e.g., `WIP`, `Quotes`). It adds a `*` to certain files to indicate a special status.
- **`Calc_Numbers.bas`**: Manages the generation of unique sequential numbers for enquiries, quotes, and jobs.
  - `Calc_Next_Number`: Determines the next available number.
  - `Confirm_Next_Number`: Calculates the next number and updates a counter file to prevent reuse, effectively "reserving" the number.
- **`Check_Dir.bas`**: A utility function `CheckDir` that ensures a directory exists, creating it if necessary, and then sets it as the current directory.
- **`Check_Updates.bas`**: Implements a background process to monitor for new files.
  - `CheckUpdates`: Runs periodically (every 5 minutes) to count files in key directories and updates the UI to notify the user of changes.
  - `Check_Files`: A helper function to count files in a directory.
- **`Delete_Sheet.bas`**: A simple function `DeleteSheet` to delete an Excel worksheet without showing a confirmation prompt to the user.
- **`GetUserNameEx.bas`**: Uses the Windows API to get the network username of the person using the application.
- **`GetValue.bas`**: A critical utility function used throughout the application to read a value from a specific cell in a **closed** Excel workbook. This improves performance by avoiding the need to open files just to retrieve small pieces of data.
- **`Module1.bas` (`Update_Search`)**: Contains a powerful subroutine to rebuild or synchronize the master `Search.xls` file. It iterates through all job, quote, and enquiry files and pulls their data into the search sheet.
- **`Module2.bas` (`Leeora`)**: Contains code to conditionally cancel the `BeforeSave` event, effectively preventing the user from saving the workbook. This seems to be a control mechanism based on username or computer name.
- **`Module3.bas` (`ExportAllModules`)**: A developer utility used to export all VBA code (forms, modules) from the Excel file into individual text files, which is how these `.bas` and `.frm` files were likely created.
- **`Open_Book.bas`**: A wrapper function that simplifies the command for opening an Excel workbook in either read-only or read-write mode.
- **`RefreshMain.bas`**: Contains the `Refresh_Main` function, which reloads the file list and data displayed on the main application form.
- **`RemoveCharacters.bas`**: Contains helper functions for string manipulation.
  - `Remove_Characters`: Removes special characters from a string, often used to create valid worksheet names.
  - `Insert_Characters`: Formats internal field names into more readable text for the UI (e.g., `Component_Description` becomes "Component Description").
- **`Search_Sync.bas`**: A password-protected utility to synchronize `Search.xls` with `Search History.xls` and to clean out old records.
- **`Very_HiddenSheet.bas`**: Contains functions to make a worksheet "very hidden" (not visible in the Excel UI) or to make it visible again.

---

## Forms (`.frm` files)

These files define the graphical user interface (GUI) and contain the code for event handling (e.g., button clicks, form loading).

- **`Main.frm`**: The main application dashboard. It provides the primary navigation for the system, allowing users to view lists of enquiries, quotes, WIP, and archived jobs. It displays summary data for selected items and contains the buttons to launch all major actions (create enquiry, make quote, accept job, etc.).
- **`FEnquiry.frm` / `FrmEnquiry.frm`**: The form for creating a new customer enquiry. It captures customer details, component information, and quantity. It saves the new enquiry into the `enquiries` folder and adds it to the `Search.xls` file.
- **`FQuote.frm`**: Used to create a formal quote from an existing enquiry. It loads the enquiry data and allows the user to add pricing and lead time. It then moves the file from the `enquiries` folder to the `Quotes` folder.
- **`FAcceptQuote.frm`**: Handles the process of converting an accepted quote into a live job. It prompts for a customer order number, generates a new job number, and moves the file from the `Quotes` folder into the `WIP` (Work In Progress) folder.
- **`FJobCard.frm`**: The form for managing a job that is currently in progress. It allows users to edit operational details, apply templates, and eventually save the completed job card, which moves the file from `WIP` to the `Archive`.
- **`FJG.frm`**: A multi-purpose form for "Jumping the Gun" or creating jobs from contract templates. It allows for the creation of a job without a preceding enquiry or quote, and can also save job details as a reusable contract template.
- **`FList.frm`**: A simple, reusable form that only contains a list box. It is used throughout the application to present a list of items (e.g., files, customers, templates) for the user to select from.
- **`fwip.frm`**: The "WIP Report" form. This powerful tool generates various printable reports from the data in `WIP.xls`. It can create reports organized by operation, by operator, or sorted by due date, customer, or job number, with different layouts for "Office" and "Workshop" use.
