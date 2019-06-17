﻿using Google.Apis.Auth.OAuth2;
using Google.Apis.Sheets.v4;
using Google.Apis.Sheets.v4.Data;
using Google.Apis.Drive.v3;
using Google.Apis.Services;
using System;
using System.Collections.Generic;
using System.IO;
using System.Threading;
using Autodesk.DesignScript.Runtime;

namespace BIMOne
{
    public static class GoogleAPI
    {
        // If modifying these scopes, delete your previously saved credentials
        // at ~/.credentials/sheets.googleapis.com-dotnet-quickstart.json
        static string[] Scopes = { SheetsService.Scope.DriveReadonly, SheetsService.Scope.Spreadsheets };
        static string ApplicationName = "BIM One Google Sheets";
        static string credentialsPath = String.Format(@"{0}\Dynamo\Dynamo Revit\2.0\packages\{1}\extra\{2}", 
            Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData), 
            System.Reflection.Assembly.GetCallingAssembly().GetName().Name, 
            "credentials.json");
        
        static UserCredential credential;

        static DriveService driveService { get => 
            // Create Google Drive API service.
            new DriveService(new BaseClientService.Initializer()
            {
                HttpClientInitializer = credential,
                ApplicationName = ApplicationName,
            });
        }
        static SheetsService sheetsService { get => 
            // Create Google Sheets API service.
            new SheetsService(new BaseClientService.Initializer()
            {
                HttpClientInitializer = credential,
                ApplicationName = ApplicationName,
            });
        }

        [IsVisibleInDynamoLibrary(false)]
        static GoogleAPI()
        {
            GetCredentials();
        }

        [IsVisibleInDynamoLibrary(false)]
        static UserCredential GetCredentials()
        {
            using (var stream =
                new FileStream(credentialsPath, FileMode.Open, FileAccess.Read))
            {
                return credential = GoogleWebAuthorizationBroker.AuthorizeAsync(
                    GoogleClientSecrets.Load(stream).Secrets,
                    Scopes,
                    "user",
                    CancellationToken.None).Result;
            }
        }

        /// <summary>
        /// Gets a list of Google Sheets present in a user's Google Drive with optional
        /// 'contains' keyword filter.
        /// </summary>
        /// <param name="filter">The text filter to use when seraching for Google Sheets on the Drive.</param>
        /// <returns>fileNames, fileIds as lists.</returns>
        /// <search>
        /// google, sheets, drive, read
        /// </search>
        [MultiReturn(new[] { "fileNames", "fileIds" })]
        public static Dictionary<string, object> GetGoogleSpreadsheets(string filter = "")
        {
            var fileNames = new List<string>();
            var fileIds= new List<string>();

            // Get all spreadsheets from drive
            // Define parameters of request.
            FilesResource.ListRequest listRequest = driveService.Files.List();
            listRequest.PageSize = 1000;
            listRequest.Fields = "nextPageToken, files(id, name)";
            listRequest.Q = String.Format("mimeType='application/vnd.google-apps.spreadsheet' and name contains '{0}'", filter);
            listRequest.OrderBy = "name";

            // List files.
            IList<Google.Apis.Drive.v3.Data.File> files = listRequest.Execute()
                .Files;

            if (files != null && files.Count > 0)
            {
                foreach (var file in files)
                {
                    fileNames.Add(file.Name);
                    fileIds.Add(file.Id);
                }
            }
            else
            {
                fileNames.Add("No sheets found");
                fileIds.Add("No sheets found");
            }

            var d = new Dictionary<string, object>();
            d.Add("fileNames", fileNames);
            d.Add("fileIds", fileIds);
            return d;
        }

        /// <summary>
        /// Appends a nested list of lists to a Google Spreadsheet
        /// </summary>
        /// <param name="spreadsheetId">The ID of the Spreadsheet (long unique identifier as string)</param>
        /// <param name="sheet">The name of the sheet within the spreadsheet as string. Ex.: Sheet1 </param>
        /// <param name="range">The range where to try and find a table, as string. Ex.: A:Z</param>
        /// <param name="data">A list of lists containing the data to append to the table in Google Sheets.</param>
        /// <param name="userInputModeRaw">If true, prevents Google sheets from auto-detecting the formatting of the data (useful for dates).</param>
        /// <returns>response</returns>
        /// <search>
        /// google, sheets, drive, write
        /// </search>
        [MultiReturn(new[] { "response" })]
        public static Dictionary<string, object> AppendDataToGoogleSheetsTable(string spreadsheetId, string sheet, string range, List<IList<object>>data, bool userInputModeRaw = false)
        {
            // Range format: SHEET:!A:F
            range = $"{sheet}!{range}";

            var valueRange = new ValueRange();
            valueRange.Values = data;

            var appendRequest = sheetsService.Spreadsheets.Values.Append(valueRange, spreadsheetId, range);
          
            if(userInputModeRaw)
            {
                appendRequest.ValueInputOption = SpreadsheetsResource.ValuesResource.AppendRequest.ValueInputOptionEnum.RAW;
            }
            else
            {
                appendRequest.ValueInputOption = SpreadsheetsResource.ValuesResource.AppendRequest.ValueInputOptionEnum.USERENTERED;
            }
            
            var appendResponse = appendRequest.Execute();

            var d = new Dictionary<string, object>();
            d.Add("response", appendResponse);
            return d;
        }

        /// <summary>
        /// Writes a nested list of lists to a Google Spreadsheet
        /// </summary>
        /// <param name="spreadsheetId">The ID of the Spreadsheet (long unique identifier as string)</param>
        /// <param name="sheet">The name of the sheet within the spreadsheet as string. Ex.: Sheet1 </param>
        /// <param name="range">The range where to write the data as string. Ex.: A:Z</param>
        /// <param name="data">A list of lists containing the data to write to Google Sheets.</param>
        /// <param name="userInputModeRaw">If true, prevents Google sheets from auto-detecting the formatting of the data (useful for dates).</param>
        /// <returns>response</returns>
        /// <search>
        /// google, sheets, drive, write
        /// </search>
        [MultiReturn(new[] { "response" })]
        public static Dictionary<string, object> WriteDataToGoogleSheets(string spreadsheetId, string sheet, string range, List<IList<object>> data, bool userInputModeRaw = false)
        {
            // Range format: SHEET:!A:F
            range = $"{sheet}!{range}";

            var valueRange = new ValueRange();
            valueRange.Values = data;

            var updateRequest = sheetsService.Spreadsheets.Values.Update(valueRange, spreadsheetId, range);

            if (userInputModeRaw)
            {
                updateRequest.ValueInputOption = SpreadsheetsResource.ValuesResource.UpdateRequest.ValueInputOptionEnum.RAW;
            }
            else
            {
                updateRequest.ValueInputOption = SpreadsheetsResource.ValuesResource.UpdateRequest.ValueInputOptionEnum.USERENTERED;
            }
            
            var updateResponse = updateRequest.Execute();

            var d = new Dictionary<string, object>();
            d.Add("response", updateResponse);
            return d;
        }

        /// <summary>
        /// Reads a specified range of a Google Spreadsheet
        /// </summary>
        /// <param name="spreadsheetId">The ID of the Spreadsheet (long unique identifier as string)</param>
        /// <param name="sheet">The name of the sheet within the spreadsheet as string. Ex.: Sheet1 </param>
        /// <param name="range">The range where to write the data as string. Ex.: A:Z</param>
        /// <param name="unformattedValues">If true, reads Google sheets as raw values.</param>
        /// <returns>data</returns>
        /// <search>
        /// google, sheets, drive, read
        /// </search>
        [MultiReturn(new[] { "data" })]
        public static Dictionary<string, object> ReadGoogleSheet(string spreadsheetId, string sheet, string range, bool unformattedValues = false)
        {
            // Range format: SHEET:!A:F
            range = $"{sheet}!{range}";

            // How values should be represented in the output.
            // The default render option is ValueRenderOption.FORMATTED_VALUE.
            SpreadsheetsResource.ValuesResource.GetRequest.ValueRenderOptionEnum valueRenderOption;
            if (unformattedValues)
            {
                valueRenderOption = SpreadsheetsResource.ValuesResource.GetRequest.ValueRenderOptionEnum.UNFORMATTEDVALUE;
            }
            else
            {
                valueRenderOption = SpreadsheetsResource.ValuesResource.GetRequest.ValueRenderOptionEnum.FORMATTEDVALUE;
            }

            // How dates, times, and durations should be represented in the output.
            // This is ignored if value_render_option is
            // FORMATTED_VALUE.
            // The default dateTime render option is [DateTimeRenderOption.SERIAL_NUMBER].
            SpreadsheetsResource.ValuesResource.GetRequest.DateTimeRenderOptionEnum dateTimeRenderOption = 
                SpreadsheetsResource.ValuesResource.GetRequest.DateTimeRenderOptionEnum.FORMATTEDSTRING; 

            SpreadsheetsResource.ValuesResource.GetRequest request = sheetsService.Spreadsheets.Values.Get(spreadsheetId, range);
            request.ValueRenderOption = valueRenderOption;
            request.DateTimeRenderOption = dateTimeRenderOption;

            var getResponse = request.Execute();

            var d = new Dictionary<string, object>();
            d.Add("data", getResponse.Values);
            return d;
        }

        /// <summary>
        /// Clears values in the specified range of cells in a Google Sheet. Only values, not formatting.
        /// </summary>
        /// <param name="spreadsheetId">The ID of the Spreadsheet (long unique identifier as string)</param>
        /// <param name="sheet">The name of the sheet within the spreadsheet as string. Ex.: Sheet1 </param>
        /// <param name="range">The range where to write the data as string. Ex.: A:Z</param>
        /// <returns>response</returns>
        /// <search>
        /// google, sheets, clear, range
        /// </search>
        [MultiReturn(new[] { "response" })]
        public static Dictionary<string, object> ClearValuesInRangeGoogleSheet(string spreadsheetId, string sheet, string range)
        {
            // Range format: SHEET:!A:F
            range = $"{sheet}!{range}";

            var requestBody = new Google.Apis.Sheets.v4.Data.ClearValuesRequest();
            SpreadsheetsResource.ValuesResource.ClearRequest request = sheetsService.Spreadsheets.Values.Clear(requestBody, spreadsheetId, range);
            var clearResponse = request.Execute();

            var d = new Dictionary<string, object>();
            d.Add("response", clearResponse.ClearedRange);
            return d;
        }

        /// <summary>
        /// Gets sheet title and ids in a Google Spreadsheet
        /// </summary>
        /// <param name="spreadsheetId">The ID of the Spreadsheet (long unique identifier as string)</param>
        /// <returns>sheetTitles</returns>
        /// <returns>sheetIds</returns>
        /// <search>
        /// google, sheets, titles, ids
        /// </search>
        [MultiReturn(new[] { "sheetTitles", "sheetIds" })]
        public static Dictionary<string, object> GetSheetsInSpreadsheet(string spreadsheetID)
        {
            SpreadsheetsResource.GetRequest request = sheetsService.Spreadsheets.Get(spreadsheetID);
            var response = request.Execute();

            var sheets = response.Sheets;
            List<string> sheetTitles = new List<string>();
            List<string> sheetIds = new List<string>();
            foreach (Sheet sheet in sheets)
            {
                sheetTitles.Add(sheet.Properties.Title);
                sheetIds.Add(sheet.Properties.SheetId.ToString());
            }

            var d = new Dictionary<string, object>();
            d.Add("sheetTitles", sheetTitles);
            d.Add("sheetIds", sheetIds);
            return d;
        }

        [MultiReturn(new[] { "spreadsheetId", "sheetUrl" })]
        public static Dictionary<string, object> CreateSpreadsheet(string spreadsheetTitle, bool openInBrowser = false)
        {
            var spreadsheet = new Spreadsheet();

            spreadsheet.Properties = new SpreadsheetProperties();
            spreadsheet.Properties.Title = spreadsheetTitle;

            SpreadsheetsResource.CreateRequest request = sheetsService.Spreadsheets.Create(spreadsheet);

            var response = request.Execute();

            if(openInBrowser)
            {
                openLinkInBrowser(response.SpreadsheetUrl);
            }
            
            var d = new Dictionary<string, object>();
            d.Add("spreadsheetId", response.SpreadsheetId);
            d.Add("sheetUrl", response.SpreadsheetUrl);
            return d;
        }

        [IsVisibleInDynamoLibrary(false)]
        static void openLinkInBrowser(string url)
        {
            System.Diagnostics.Process.Start(url);
        }
    }
}