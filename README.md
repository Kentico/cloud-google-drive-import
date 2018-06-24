
# cloud-google-drive-import
This console application can import files from your personal Google Drive account into Kentico Cloud as content items or assets. It can be run without user interaction via command line for regular, scripted imports, or you can use a GUI to provide options.

### Setup

Clone the repository and build it in Visual Studio.

Before you begin, you must authorize this application to access your Google Drive account:
1. Visit [Google's developer console](https://console.developers.google.com/start/api?id=drive) to create or select a project in the Google Developers Console and automatically turn on the API. Click **Continue**, then **Go to credentials**.
2. On the **Add credentials to your project** page, click the **Cancel** button.
3. At the top of the page, select the **OAuth consent screen** tab. Select an **Email address**, enter any **Product name**, and click the **Save** button.
4. Select the **Credentials tab**, click the **Create credentials** button and select **OAuth client ID**.
5. Select the application type **Other**, enter the name "GoogleDriveImport", and click the **Create** button.
6. Click **OK** to dismiss the resulting dialog.
7. Click the **Download JSON** button to the right of the client ID.
8. Move this file to the directory containing DriveImportCore.dll and rename it **client_secret.json**.

The first time you run the application, it will prompt you to authorize access. Once authorized, this will not happen again.

You must also modify the GoogleDriveImport.json file to include your Kentico Cloud **Project ID**, **Preview API key**, and **Content Management API Key**. These keys can be found in **the API Keys** page on [https://app.kenticocloud.com](https://app.kenticocloud.com).

### Usage

In the directory containing the GoogleDriveImport.dll, open a command prompt and type `dotnet GoogleDriveImport.dll`. If you specify no arguments, a GUI will appear and walk you through the process. You may also specify the following arguments to run the program “silently” from command line:

-	`-s <filename>` or `-source <filename>`
	The name of the file to import, as it appears in Google Drive.
-	`-d <filename>` or `-dir <filename>`
	The name of the folder to import contents from (nullifies -source parameter if present). 
-	`-t <codename>` or `-type <codename>`
	The code name of the content type in Kentico Cloud to create for new items.
-	`-e <codename>` or `-element <codename>`
	The code name of the element in Kentico Cloud in which the file's contents will be inserted.
-	`-u` or `-update`
	If passed, content items with the same code name as the imported file(s) will be updated with the new content. If not passed, new content items will always be created.

Currently, the importing of a directory’s contents is only supported when supplying parameters from command line, not using the GUI.

### Supported file types
The following file types will be imported into Kentico Cloud as assets:
-	.jpg/.jpeg
-	.png
-	.gif
-	.svg

These file types will be imported as content items:
- Google Doc
-	Google Sheet
-	.txt
-	.xlsx

### Importing spreadsheets
If a spreadsheet is imported, each row of the spreadsheet will be imported as a separate content item. The first row of the spreadsheet should contain the codenames of the elements for the content type.

 **Important**: the first row of the spreadsheet must contain one header with the text "Name." Other cells under this column should contain the code name of the content item to be created/updated.
