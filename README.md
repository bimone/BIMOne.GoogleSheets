# BIMOne.GoogleSheets
A set of Dynamo nodes to facilitate interactions with Google Sheets.

## A word of caution
__Be very careful what you pass into these nodes.__ If you pass a list of spreadsheet IDs into certain nodes, Dynamo will run the node once per item in that list. For example, if you were to use the `AppendDataToGoogleSheetTable` node and pass in a list of spreadsheet IDs, Dynamo will append your data once for __*every*__ Google Sheet in that list.

Also, be careful not to confuse the concept of a Google Sheet (the entire spreadsheet) and a "sheet" within that spreadsheet.

## Installation & Config
1. Get the package installed from https://dynamopackages.com/
2. Create Google API credentials for your use:
    1. Create a [Google APIs Console project](console.cloud.google.com).
    2. [Enable the Google Drive API and Google Sheets API](https://support.google.com/googleapi/answer/6158841?hl=en&ref_topic=7013279) on this project.
    3. Create [Google API credentials](https://console.developers.google.com/apis/credentials). Must have at least the following [scopes enabled](https://support.google.com/a/answer/162106?hl=en):
        - `../auth/spreadsheets`
        - `../auth/drive`
    4. Download the credentials file (JSON):
    
        ![download credentials file](readme_images/download_credentials.png)
    5. Rename the file to `credentials.json`
    6. Place the credentials.json file in the `/extra` folder where you installed the package. For Dynamo 2.0 with locally installed packages, this would typically be `%appdata%\Dynamo\Dynamo Revit\2.0\packages\BIMOneGoogleAPI\extra`
3. Keep an eye on the [Google Sheets API usage limits](https://developers.google.com/sheets/api/limits).

## Usage
The package provides the following nodes:

| Node name  | Use |
| --- | --- |
| GetGoogleSheetsInGoogleDrive | Get a list of all the Google Sheets that you can see in you Google Drive. Optionally, pass a string to the filter parameter to narrow down the list of results. |
| GetSheetsInGoogleSheet | Get a list of the sheets (aka those "tabs" at the bottom) within a Google Sheet |
| ReadGoogleSheet | |
| ReadGoogleSheetMultipleRanges | |
| CreateNewGoogleSheet | |
| CreateNewSheetWithinGoogleSheet | |
| BatchAppendDataToGoogleSheet | |
| AppendDataToGoogleSheet | |
| WriteDataToGoogleSheet | |
| ClearValuesInRangeGoogleSheet | |
| DeleteSheetByIdWithinGoogleSheet | |
| DeleteSheetByTitleWithinGoogleSheet | |


## Need helping setting it up?
[Contact us](https://bimone.com/en/ContactUs), we will be glad to help!
