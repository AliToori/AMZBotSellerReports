
<h1>AMZBot</h1>
AMZBot is a bot to retrieve, update and upload amazon seller and vendor central reports to a Google Sheet to be easily accessed and analyzed.
It has also an option for saving local copies of the reports in the local directory called Reports.

<h2>Main Tasks</h2>
Automates the following reports from multiple amazon seller accounts
 - Business Reports
 - Advertising Reports
 - Fulfillment Reports
 - Vendor Reports
 - Vendor Promotional Reports

<h1>Features</h1>

- Signs in to amazon account using credentials given in input CSV file, Accounts.csv.
- Saves sessions and cookies locally so the user will be auto-logged-in next time and onward.
- Navigates business reports.
- Filters business reports by ASIN (Detail Page Sales and Traffic by Child Item).
- Filters business reports by From Date and To Date given in the input CSV as StartDate EndDate.
- Daily Reports: It retrieves reports for every single day starting from StartDate to EndDate.
- Weekly Reports: It retrieves reports for a week starting from StartDate to EndDate.
- Bi-Weekly Reports: It retrieves reports for 14 days starting from StartDate to EndDate.

Usage
-
1. Extract and navigate to the AMZBot => AMZRes directory.
2. Open and edit Accounts.csv file as following.
3. Put your amazon login email ID under Email. 
4. Put your amazon login password in under Password. 
5. Put your Google SpreadSheet name under SpreadSheetName. 
6. Put your WorkSheet name under WorkSheetName. 
8. For weekly day-wise reports retrieval, put a date range of 7 days e.g. 06/05/2021 and 06/12/2021.  
9. For monthly day-wise reports retrieval, put a date range of 30 days e.g. 06/05/2021 and 07/05/2021. 
10. Finally, Launch the bot by double clicking the AMZBot.exe file in main AMZBot directory. Follow input parameters for detailed inputs.

Input Parameters: Accounts.xlsx
-
<b>AccountNo:</b> A serial number used by the Bot

<b>Client:</b> Google Sheet's URL which the data to be uploaded to.

<b>Brand:</b> A brand's name which you want to retrieve reports for.

<b>Location:</b> Location/Country of the brand which you want to retrieve reports for.

<b>ReportType:</b>: Report type you want to retrieve either "Business Reports" or "Advertising Reports".

<b>Business Report Type:</b> for switching among the following business report types.
- Detail Page Sales and Traffic
- Detail Page Sales and Traffic by Parent Item
- Detail Page Sales and Traffic by Child Item

<b>Ad Campaign Type:</b> for switching among the following campaign types of advertising reports
- Sponsored Products
- Sponsored Brands
- Sponsored Brands Video
- Sponsored Display

<b>Ad Report Type:</b> for switching among the following report types in advertising reports.
- Search term
- Targeting
- Advertised product
- Campaign
- Budget
- Placement
- Purchased product
- Performance Over Time
- Search Term Impression Share
- Keyword
- Keyword placement
- Campaign placement
- Category benchmark

<b>Email:</b> Email ID of your Amazon account for signing-in.

<b>Password:</b> Password of your Amazon account for signing-in.

<b>SpreadSheetURL:</b> Google SpreadSheet's URL which the data to be uploaded to.

<b>SpreadSheetName:</b> Google SpreadSheet's name which the data to be uploaded to.

<b>WorkSheetName:</b> SpreadSheet's WorkSheet name which the data to be uploaded to.

- <b>Note:</b> If you leave a WorkSheetName as empty, the bot will upload data to the first WorkSheet by default. 
 
<b>UpdateType:</b> "Upload" or "Append" the reports to the Google Sheet. Upload will overwrite all the previous reports.

<b>SaveLocalCopy:</b> "Yes" or "No". If set to Yes, will save a copy of a report in local Reports directory.

<b>StartDate:</b> The start date which reports to be retrieved from. 

<b>EndDate:</b> The end date which reports to be retrieved from, starting from the StartDate.

<b>NOTE:</b> DO NOT DELETE THE "client-secret_victor.json" FILE IT IS BEING USED FOR GOOGLE-SHEET AUTHENTICATION.

Suggestion
-
You can change your Google Sheet as far as it is from the same google account as the client_secret.