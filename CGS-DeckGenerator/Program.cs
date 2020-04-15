using Google.Apis.Auth.OAuth2;
using Google.Apis.Services;
using Google.Apis.Sheets.v4;
using Google.Apis.Util.Store;
using System;
using System.IO;
using System.Reflection;
using System.Threading;


namespace CGS_DeckGenerator
{
    class Program
    {
        static void Main(string[] args)
        {
            try
            {
                //All the decks
                DeckItem resourceDeck = new DeckItem();
                DeckItem technologyDeck = new DeckItem();
                DeckItem eventDeck = new DeckItem();

                //Path to project
                string projectPath = AssemblyDirectory.Replace(@"\bin", "").Replace(@"\Debug", "") + '\\';

                //Read Google SpreadSheet
                if (ReadSpreadSheet(ref resourceDeck, ref technologyDeck, ref eventDeck, projectPath))
                {
                    //Attempt to draw the cards
                    Console.WriteLine("Finished reading spreadsheet. Attempting to draw deck");

                    resourceDeck.DrawDeck(projectPath, "Resource Deck");
                    technologyDeck.DrawDeck(projectPath, "Technology Deck");
                    eventDeck.DrawDeck(projectPath, "Event Deck");

                    Console.WriteLine("Finished drawing all decks");
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine("Unexpected error: " + ex.ToString());
                Console.WriteLine("Press any key to finish...");
                Console.Read();
            }
        }

        static private bool ReadSpreadSheet(ref DeckItem rscDeck, ref DeckItem rhDeck, ref DeckItem evDeck, string projectPath)
        {
            bool readSpreadSheet = false;

            try
            {
                UserCredential credential;
                string[] Scopes = { SheetsService.Scope.SpreadsheetsReadonly };

                string ApplicationName = "Ant Game Cards";

                using (FileStream stream = new FileStream(projectPath + @"credentials.json", FileMode.Open, FileAccess.Read))
                {
                    // The file token.json stores the user's access and refresh tokens, and is created
                    // automatically when the authorization flow completes for the first time.
                    string credPath = projectPath + @"token.json";
                    credential = GoogleWebAuthorizationBroker.AuthorizeAsync(
                        GoogleClientSecrets.Load(stream).Secrets, Scopes, "user",
                        CancellationToken.None, new FileDataStore(credPath, true)).Result;

                    Console.WriteLine("Credential file saved to: " + credPath);
                }

                // Create Google Sheets API service.
                SheetsService service = new SheetsService(new BaseClientService.Initializer()
                {
                    HttpClientInitializer = credential, ApplicationName = ApplicationName,
                });

                // Define request parameters.
                string spreadsheetId = "12kuqGt0t8tgm2-dP2C4qKVbbUm0R-RZ7uu75JmcAnuE";

                string resourceRange = "Resource!A2:D";
                string technologyRange = "Technology!A2:D";
                string eventRange = "Event!A2:D"; //Only first three columns, there is no grouping

                //Gather the information and fill the three decks
                rscDeck.GatherSpreadSheetData(ref service, spreadsheetId, resourceRange);
                rhDeck.GatherSpreadSheetData(ref service, spreadsheetId, technologyRange);
                evDeck.GatherSpreadSheetData(ref service, spreadsheetId, eventRange);

                //We haven't crashed so assume we're good
                readSpreadSheet = true;
            }
            catch (Exception ex)
            {
                throw new Exception("Issue with reading the spreadsheet: " + ex.ToString());
            }

            return readSpreadSheet;
        }

        public static string AssemblyDirectory
        {
            get
            {
                string codeBase = Assembly.GetExecutingAssembly().CodeBase;
                UriBuilder uri = new UriBuilder(codeBase);
                string path = Uri.UnescapeDataString(uri.Path);
                return Path.GetDirectoryName(path);
            }
        }
    }
}
