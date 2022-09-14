using Google.Apis.Auth.OAuth2;
using Google.Apis.Sheets.v4;
using Google.Apis.Sheets.v4.Data;
using Google.Apis.Services;
using Google.Apis.Util.Store;
using Google.Apis.Drive.v3;


using System;
using System.Collections.Generic;
using System.IO;
using System.Threading;
using Newtonsoft.Json;

namespace SheetsQuickstart
{
    // Class to demonstrate the use of Sheets list values API
    class Program
    {
        /* Global instance of the scopes required by this quickstart.
         If modifying these scopes, delete your previously saved token.json/ folder. */
        static string[] Scopes = {  DriveService.Scope.Drive, SheetsService.Scope.Spreadsheets };
        static string ApplicationName = "MOLOLO ";


        private static void CreateFolder(string folderName, DriveService service)
        {
            
            var fileMetadata = new Google.Apis.Drive.v3.Data.File()
            {
                Name = folderName,
                MimeType = "application/vnd.google-apps.folder"
            };
            
            var request = service.Files.Create(fileMetadata);
            
            request.Fields = "id";
            var file = request.Execute();
            Console.WriteLine("Folder ID: " + file.Id);

        }


        public static IEnumerable<Google.Apis.Drive.v3.Data.File> GetFiles(string folder, DriveService ds)
        {
            var service = ds;

            var fileList = service.Files.List();

            //fileList.Q = $"mimeType!='application/vnd.google-apps.folder' and '{folder}' in parents ";
            //fileList.Q = $"mimeType=='{folder}' in parents ";
            fileList.Fields = "nextPageToken, files(id, name, size, mimeType)";

            var result = new List<Google.Apis.Drive.v3.Data.File>();
            string pageToken = null;
            do
            {
                fileList.PageToken = pageToken;
                var filesResult = fileList.Execute();
                var files = filesResult.Files;
                pageToken = filesResult.NextPageToken;
                
                result.AddRange(files);
            } while (pageToken != null);


    return result;
        }


       

        private static void UpdatGoogleSheetinBatch(IList<IList<Object>> values, string spreadsheetId, string newRange, SheetsService service)
        {
            SpreadsheetsResource.ValuesResource.AppendRequest request =
               service.Spreadsheets.Values.Append(new ValueRange() { Values = values }, spreadsheetId, newRange);
            request.InsertDataOption = SpreadsheetsResource.ValuesResource.AppendRequest.InsertDataOptionEnum.INSERTROWS;
            request.ValueInputOption = SpreadsheetsResource.ValuesResource.AppendRequest.ValueInputOptionEnum.RAW;
            var response = request.Execute();
        }


        static void Main(string[] args)
        {
            try
            {
                UserCredential credential;
                GoogleCredential credentialD = null;
                // Load client secrets.
                using (var stream = new FileStream(".\\token.json", FileMode.Open, FileAccess.Read))
                {
                    /* The file token.json stores the user's access and refresh tokens, and is created
                     automatically when the authorization flow completes for the first time. */
                    string credPath = ".\\token2.json";
                    try
                    {
                        credential = GoogleWebAuthorizationBroker.AuthorizeAsync(GoogleClientSecrets.FromStream(stream).Secrets, Scopes, "user", CancellationToken.None, new FileDataStore(credPath, true)).Result;
                    }
                    catch (Exception e) { Console.WriteLine(e.ToString()); credential = null; }

                    /*try
                    {
                     credentialD = GoogleCredential.GetApplicationDefault().CreateScoped(DriveService.Scope.Drive); 
                    }
                    catch (Exception e) { Console.WriteLine(e.ToString()); credential = null; }*/
                    Console.WriteLine("Credential file saved to: " + credPath);
                }

                // Create Google Sheets API service.
                var service = new SheetsService(new BaseClientService.Initializer
                {
                    HttpClientInitializer = credential,
                    ApplicationName = ApplicationName
                });

                var serviceDrive = new DriveService(new BaseClientService.Initializer
                {
                    HttpClientInitializer = credential,
                    ApplicationName = ApplicationName
                });


                // Define request parameters.
                String spreadsheetId = "1BxiMVs0XRA5nFMdKvBdBZjgmUUqptlbs74OgvE2upms";
                String range = "Class Data!A2:E";
                SpreadsheetsResource.ValuesResource.GetRequest request = service.Spreadsheets.Values.Get(spreadsheetId, range);

                // Prints the names and majors of students in a sample spreadsheet:
                // https://docs.google.com/spreadsheets/d/1BxiMVs0XRA5nFMdKvBdBZjgmUUqptlbs74OgvE2upms/edit
                ValueRange response = request.Execute();
                IList<IList<Object>> values = response.Values;
                if (values == null || values.Count == 0)
                {
                    Console.WriteLine("No data found.");
                    return;
                }
                Console.WriteLine("Name, Major");
                foreach (var row in values)
                {
                    // Print columns A and E, which correspond to indices 0 and 4.
                    Console.WriteLine("{0}, {1}", row[0], row[4]);
                }
                CreateFolder("KOKO", serviceDrive);
                Console.WriteLine("LISTANDO FICHEROS");
                foreach (var ffile in GetFiles("root", serviceDrive))
                {
                    Console.WriteLine(ffile.Name + " <<<-ID->>>" + ffile.Id + "------------" + ffile.Kind + "----" + ffile.MimeType);
                    if (ffile.MimeType.Contains("vnd.google-apps.spreadsheet"))
                    {

                        // Specifying Column Range for reading...
                        range = $"MOLOLO!A:E";
                        var valueRange = new ValueRange();
                        // Data for another Student...
                        var oblist = new List<object>() { "MOLOLO1", "EY2", "King ", "Kong", "98" };
                        valueRange.Values = new List<IList<object>> { oblist };
                        // Append the above record...
                        var appendRequest = service.Spreadsheets.Values.Append(valueRange, ffile.Id, range);
                        appendRequest.ValueInputOption = SpreadsheetsResource.ValuesResource.AppendRequest.ValueInputOptionEnum.USERENTERED;
                        var appendReponse = appendRequest.Execute();


                        SpreadsheetsResource.GetRequest request2 = service.Spreadsheets.Get(ffile.Id);
                        Spreadsheet response2 = request2.Execute();
                        Console.WriteLine(response2.Sheets[0].ToString());
                        Console.WriteLine(JsonConvert.SerializeObject(response2));


                        string sheetName = string.Format("{0} {1}", DateTime.Now.Month, DateTime.Now.Day);
                        var addSheetRequest = new AddSheetRequest();
                        addSheetRequest.Properties = new SheetProperties();
                        addSheetRequest.Properties.Title = "MOLOLO";
                        BatchUpdateSpreadsheetRequest batchUpdateSpreadsheetRequest = new BatchUpdateSpreadsheetRequest();
                        batchUpdateSpreadsheetRequest.Requests = new List<Request>();
                        batchUpdateSpreadsheetRequest.Requests.Add(new Request
                        {
                            AddSheet = addSheetRequest
                        });

                        var batchUpdateRequest = service.Spreadsheets.BatchUpdate(batchUpdateSpreadsheetRequest, ffile.Id);
                        try
                        {
                            batchUpdateRequest.Execute();
                        } catch (Exception e ) { Console.WriteLine("YA EXISTE HOJA"); }


                    }
                }
            }
            catch (FileNotFoundException e)
            {
                Console.WriteLine(e.Message);
            }
        }
    }
}