# Feed-Generator

This desktop application was developed as an initiative during my time as an intern in Autodesk. It is meant for SAPConcur, where during the creation of expense types for card payments, a card feed, a .txt file with details of the expense, is to be submitted.

For the purposes of testing, manually creating the feeds was time consuming. Thus, I built this app on Python using PyQt5 and pandas library to allow for easy creation of card feeds.

When creating feeds with the .xlsx file, it allows the user to save the relevant information ad generate feed for multiple employees at the same time.

Here is an example of how a card feed looks:

![Feed](/images/Feed.png)


## Functionalities

- Creation of card feed by inputting data into the GUI

	1. Click on the "Create New Feed" button in the main menu.
		
		<img src="/images/ClickCreateNew.png" alt="Click Create New button" style="height: 400px; width:400px;"/>
	
	2. Fill in the relevant details in the following screen:

		<img src="/images/CreateMenu.png" alt="Create Menu Image"/>

	3. Click on the "Generate" button and the feed will be created in the Generated_Files folder.
	
		<img src="/images/GeneratedFolder.png" alt="Generated Folder"/>


- Creation of card feed by inputting data into excel file

	1. During the first run of the application, an excel file named feed.xlsx with the relevant columns will generated in the Excel_Feed folder

		<img src="/images/ExcelFolder.png" alt="Excel Folder"/>

	2. The user can then fill in the relevant details and quantity required into the excel

		<img src="/images/Excel.png" alt="Excel file"/>

	3. Once done, save the excel and close it, before clicking the "Create from Excel" option in the application to generate the feed.

		<img src="/images/ClickCreateExcel.png" alt="Click Create from Excel button" style="height: 400px; width:400px;"/>

	4. The feed(s) will be created in the Generated_Files folder.


## File Structure

![File Structure](/images/FileStructure.png)

The Excel_Feed folder is created during the first run of the app, and within it the feed.xlsx file will also be created.

The Generated_Files folder is created during the first creation of any feed, and all feeds created are stored in this folder.


## Getting Started

To try out the application, simple download the .exe file from Releases and run it from your computer.
