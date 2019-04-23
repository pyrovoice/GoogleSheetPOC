using Google.Apis.Auth.OAuth2;
using Google.Apis.Sheets.v4;
using Google.Apis.Sheets.v4.Data;
using Google.Apis.Services;
using Google.Apis.Util.Store;
using System;
using System.Collections.Generic;
using System.IO;
using System.Threading;

namespace GoogleSheetPOC
{
    class Program
    {
        // If modifying these scopes, delete your previously saved credentials
        // at ~/.credentials/sheets.googleapis.com-dotnet-quickstart.json
        static string[] Scopes = { SheetsService.Scope.SpreadsheetsReadonly };
        static string ApplicationName = "Google Sheets API .NET Quickstart";

        static void Main(string[] args)
        {
            UserCredential credential;

            using (var stream =
                new FileStream("Ressources\\credentials.json", FileMode.Open, FileAccess.Read))
            {
                // The file token.json stores the user's access and refresh tokens, and is created
                // automatically when the authorization flow completes for the first time.
                string credPath = "token.json";
                credential = GoogleWebAuthorizationBroker.AuthorizeAsync(
                    GoogleClientSecrets.Load(stream).Secrets,
                    Scopes,
                    "user",
                    CancellationToken.None,
                    new FileDataStore(credPath, true)).Result;
                Console.WriteLine("Credential file saved to: " + credPath);
            }

            // Create Google Sheets API service.
            var service = new SheetsService(new BaseClientService.Initializer()
            {
                HttpClientInitializer = credential,
                ApplicationName = ApplicationName,
            });

            // Define request parameters.
            String spreadsheetId = "1iRaiRkFeWnsx8602r_SrJJsf-XS67emTXEDutychebU";
            Spreadsheet spreadsheet = service.Spreadsheets.Get(spreadsheetId).Execute();
            foreach (Sheet sheet in spreadsheet.Sheets)
            {
                Console.WriteLine(sheet.Properties.Title);
                String range = sheet.Properties.Title + "!A2:G";
                SpreadsheetsResource.ValuesResource.GetRequest request =
                        service.Spreadsheets.Values.Get(spreadsheetId, range);
                IList<IList<Object>> values = request.Execute().Values;

                if (values != null && values.Count > 0)
                {
                    Console.WriteLine("Name, Cost, Type, Subtype, Text, Stats, Art");
                    foreach (var row in values)
                    {
                        // Print columns A and E, which correspond to indices 0 and 4.
                        for (int i = 0; i < row.Count; i++)
                        {
                            Console.Write("{0}, ", row[i]);
                        }
                        Console.WriteLine();
                    }
                }
                Console.WriteLine();
            }

            Console.Read();
        }
    }
}