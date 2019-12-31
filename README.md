# mergeutility
## Description
Utility that provides Mail Merge functionality to Google Sheets and Docs

## Setup
To set up this utility for a mail merge document, you'll need to:
* Create the template document
* Create the merge data
* Add the utility script to the merge data sheet

Once those steps are done, you can execute the utility as many times as you need. To set up another mail merge document, repeat the above three steps.

### Create the template document
Open Google Drive
![Google Drive main page](images/Template1.png]

Create a new Google Doc (`+ New` :arrow_right: `Google Docs` :arrow_right: `Blank document`)
![Creating a new Google Doc](images/Template2.png]

Enter the body of your template, using `<<` and `>>` to indicate variables that should be replaced as part of your mailmerge.
![Draft your template](images/Template3.png]

Give your document a name.

*NOTE*: The name must _not _match any other document in any folder within your Drive. This is super-important, otherwise the merge utility won't work.

Also, make sure you remember the name because it'll be used later.
![Name your document, and remember the name](images/Template4.png)

Done with the template

### Create the merge data
In Google Drive, create a new Google Sheet (`+ New` :arrow_right: `Google Sheets` :arrow_right: `Blank spreadsheet`)
![Creating a new Google Sheet](images/MergeData1.png)
In the first row, enter each of the names of the variables you used in your template doc
![Variable names as column headers](images/MergeData2.png)
For each copy of the letter you want to create, enter a new row in the sheet
![Populate your merge data](images/MergeData3.png)
Give your spreadsheet a name. Any name will do.
![Name your spreadsheet](images/MergeData4.png)
Name the tab to match the template document name exactly.
![Name your tab](images/MergeData5.png)
Done with the merge data

### Add the utility script to the merge data sheet
On the merge data spreadsheet, open the `Tools` menu and select the `Script Editor` option
![Open the script editor](images/UtilitySetup1.png)
The script editor should open a file called `Code.gs` in a new tab (or window) of your browser, with an empty function called `myFunction`
![Script editor](images/UtilitySetup2.png)
Replace all the contents of the script editor with contents of the utility script (all the code) from X:
![Utility script code into the project](images/UtilitySetup3.png)
Select `File` :arrow_right: `Save` and enter a project name
![Name your project](images/UtilitySetup4.png)
Done with setup

### To execute the script
The first time you want to execute the script after setting it up, you'll need to close all the documents you created (template doc and merge data spreadsheet).
Open the merge data spreadsheet. If this is the first time opening the spreadsheet after setting up the utility, it'll take some time to open. Once it's open, you should see a new menu option for `Merge`, after the `Help` menu
![New menu](images/Execution1.png)
Opening the Merge menu should give you one option called `Merge`
![New menu option](images/Execution2.png)
Selecting the `Merge` menu option will execute the script. If this is your first time running the script for this spreadsheet, you will get a popup for Authorization Required.
![Auth required](images/Execution3.png)
Click Continue. Google will ask you to choose an account. Select your currently logged in account
![Choose account](images/Execution4.png)
You'll get a warning message indicating that the script is not verified. Since it's being installed _by you_ on your account, and you have access to the source code, you should feel comfortable with proceeding. Select the `Advanced` link in the lower left corner of the window.
![Warning for unverified app](images/Execution5.png)
Read the cautionary verbiage at the bottom of the screen and select the `Go to Merge` link at the bottom
![Advanced warning](images/Execution6.png)
You'll get a screen presenting the permissions that the script is requesting on your account. Confirm that your account is at the top of the screen, and click the `Allow` button.
![Allow permissions](images/Execution7.png)
Once you allow permissions, the script should execute *for the tab that you're currently on* in the spreadsheet, using the same name to look up the template document. The script will display a message when it's complete, with the name of the output file it generated as a result of the mail merge. The output file name will be the same as the template, with the word `output` and a date/timestamp value appended. For example, `Happy Birthday Form Letter-output-20191231-220602-762` which represents the execution of the script against the template `Happy Birthday Form Letter` at the time of writing - on 2019/12/31 (New Year's Eve) at 22:06:02.762 UTC (2.762 seconds after 5:02pm Eastern Daylight Time)
![Script execution complete](images/Execution8.png)
When you navigate back to Drive, you should see the new file
![Output file in Drive listing](images/Execution9.png)
When you open the file, you should see the results of the mail merge. In my example, I see an 8-page document, with one birthday note to each recipient on their birthday
![Output file](images/Execution10.png)

