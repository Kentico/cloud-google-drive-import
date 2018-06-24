using Google.Apis.Services;
using Google.Apis.Util.Store;
using KenticoCloud.Delivery;
using System;
using System.IO;
using System.Net;
using System.Threading;
using Newtonsoft.Json.Linq;
using Google.Apis.Drive.v3;
using Google.Apis.Auth.OAuth2;
using KenticoCloud.ContentManagement;
using System.Collections.Generic;
using System.Linq;

namespace DriveImportCore
{
    class Program
    {
        const string AppName = "GoogleDriveImport";
        const string ConfigFile = "GoogleDriveImport.json";
        const string ClientSecretFile = "client_secret.json";

        static DriveHelper driveHelper;

        static void Main(string[] args)
        {
            if (Init())
            {
                if (args.Length == 0) {
                    // Open with GUI
                    Run();
                }
                else
                {
                    // Run autoimport with passed args
                    CloudHelper.AutoImport(args, driveHelper);
                }
            }
        }

        /// <summary>
        /// Displays a list of files from Google Drive and allows the user to select one
        /// </summary>
        /// <param name="files"></param>
        /// <returns></returns>
        public static Google.Apis.Drive.v3.Data.File VerifySourceFile(IList<Google.Apis.Drive.v3.Data.File> files)
        {
            Program.WriteColoredText("Files:", ConsoleColor.Green);
            for(var i = 0; i < files.Count; i++)
            {
                Console.WriteLine($"{i+1}. {files[i].Name}");
            }

            Program.WriteColoredText("Which file would you like to import? [Enter a number]", ConsoleColor.Green);
            var num = Console.ReadLine();
            if (int.TryParse(num, out int index) && index >= 1 && index <= files.Count)
            {
                index--; //Subtract one since index is zero-based
                return files[index];
            }
            else
            {
                Program.ShowError("Unrecognized or out-of-bounds number. Please press any key to try again.");
                Console.ReadKey();
                return VerifySourceFile(files);
            }
        }

        /// <summary>
        /// Displays a list of content types from Kentico Cloud and allows the user to select one
        /// </summary>
        /// <param name="types"></param>
        /// <returns></returns>
        public static ContentType VerifyContentType(DeliveryTypeListingResponse types)
        {
            Console.WriteLine();
            Program.WriteColoredText("Which content type would you like to create? [Enter a number]", ConsoleColor.Green);

            for (int i = 0; i < types.Types.Count; i++)
            {
                Console.WriteLine($"{i+1}. {types.Types[i].System.Name}");
            }

            int index = 0;
            var num = Console.ReadLine();
            if (int.TryParse(num, out index) && index >= 1 && index <= types.Types.Count)
            {
                index--; // Subtract one since index is zero-based

                // Validate that the type has supported elements
                var elements = CloudHelper.FilterElements(types.Types[index].Elements);
                if (elements.Count == 0)
                {
                    Program.WriteColoredText("Chosen content type has no Text or Rich Text elements, please choose another.", ConsoleColor.Green);
                    Console.ReadKey();
                    return VerifyContentType(types);
                }

                return types.Types[index];
            }
            else
            {
                Program.ShowError("Unrecognized or out-of-bounds number. Please press any key to try again.");
                Console.ReadKey();
                return VerifyContentType(types);
            }
        }

        /// <summary>
        /// Displays a list of elements from the passed content type and allows the user to select one. Elements must be Text or Rich Text.
        /// </summary>
        /// <param name="contenttype"></param>
        /// <returns></returns>
        public static ContentElement VerifyContentElement(ContentType contenttype)
        {
            var elements = CloudHelper.FilterElements(contenttype.Elements);

            Console.WriteLine();
            Program.WriteColoredText("Which element should the imported text be inserted into? [Enter a number]", ConsoleColor.Green);
            for(int i = 0; i<elements.Count; i++)
            {
                Console.WriteLine($"{i+1}. {elements[i].Key} ({elements[i].Value.Type})");
            }
            int index = 0;
            var num = Console.ReadLine();
            if (int.TryParse(num, out index) && index >= 1 && index <= elements.Count)
            {
                index--; //Subtract one since index is zero-based
                return elements[index].Value;
            }
            else
            {
                Program.ShowError("Unrecognized or out-of-bounds number. Please press any key to try again.");
                Console.ReadKey();
                return VerifyContentElement(contenttype);
            }

        }

        public static void Run()
        {
            ContentType targetType = null;
            ContentElement targetElement = null;
            Google.Apis.Drive.v3.Data.File sourceFile = null;
            bool update = false;
            bool hasError = false;

            var files = driveHelper.ListFiles();
            if (files != null && files.Count > 0)
            {
                 sourceFile = VerifySourceFile(files);
            }
            else
            {
                Program.ShowError("No files found.");
                Console.ReadKey();
                return;
            }
            
            if (!DriveHelper.IsAsset(sourceFile))
            {
                var types = CloudHelper.ListContentTypes();
                targetType = VerifyContentType(types);

                // If imported file is not a spreadsheet, get target element
                if (!DriveHelper.IsSpreadsheet(sourceFile))
                {
                    targetElement = VerifyContentElement(targetType);
                }

                update = AskUpdate();
            }

            // Verify data and begin import
            Console.WriteLine();
            if (!CanImport(sourceFile, targetElement, targetType))
            {
                ShowError("Something went wrong.. Please press any key to close.");
                Console.ReadKey();
                return;
            }
            else
            {
                try
                {
                    var stream = driveHelper.DownloadFile(sourceFile);
                    if (stream != null) CloudHelper.BeginImport(stream, sourceFile, targetElement, targetType, update);
                }
                catch(Exception e)
                {
                    hasError = true;
                    ShowError(e.Message);
                }
                finally
                {
                    Console.WriteLine(hasError ? "Import completed with error(s), please check error messages" : "Import complete");
                    // Ask if user wants to run import again
                    Console.WriteLine("Do you want to perform another import? [Enter = Yes, Any other key = Quit]");
                    if (Console.ReadKey().Key == ConsoleKey.Enter)
                    {
                        Run();
                    }
                }
            }
        }

        private class KenticoCloudConfig
        {
            public string ApiKey;
            public string ProjectID;
            public string PreviewKey;

            public KenticoCloudConfig(string apiKey, string projectId, string previewKey)
            {
                ApiKey = apiKey;
                ProjectID = projectId;
                PreviewKey = previewKey;
            }
        }

        private static KenticoCloudConfig LoadKenticoCloudConfiguration(string jsonFile)
        {
            JObject settings = JObject.Parse(File.ReadAllText(jsonFile));
            string apiKey = settings.SelectToken("API_KEY").ToString();
            string projectID = settings.SelectToken("PROJECT_ID").ToString();
            string previewKey = settings.SelectToken("PREVIEW_KEY").ToString();

            return new KenticoCloudConfig(apiKey, projectID, previewKey);
        }

        private static void InitKenticoCloud(KenticoCloudConfig config)
        {
            ContentManagementOptions cmoptions = new ContentManagementOptions
            {
                ProjectId = config.ProjectID,
                ApiKey = config.ApiKey
            };

            CloudHelper.Init(
                new DeliveryClient(config.ProjectID, config.PreviewKey),
                new ContentManagementClient(cmoptions));
        }

        private static UserCredential LoadGoogleDriveCredentials()
        {
            UserCredential result;
            using (var stream = new FileStream(ClientSecretFile, FileMode.Open, FileAccess.Read))
            {
                string credPath = Environment.GetFolderPath(Environment.SpecialFolder.Personal);
                credPath = Path.Combine(credPath, ".credentials/kcdriveimport.json");
                result = GoogleWebAuthorizationBroker.AuthorizeAsync(
                    GoogleClientSecrets.Load(stream).Secrets,
                    new string[] { DriveService.Scope.DriveReadonly },
                    "user",
                    CancellationToken.None,
                    new FileDataStore(credPath, true)).Result;
            }
            return result;
        }

        static bool Init()
        {
            Console.ForegroundColor = ConsoleColor.Gray;
            try
            {
                KenticoCloudConfig cloudConfig = LoadKenticoCloudConfiguration(ConfigFile);
                InitKenticoCloud(cloudConfig);

                UserCredential credential = LoadGoogleDriveCredentials();
                driveHelper = new DriveHelper(credential, AppName);
            }
            catch (Exception e)
            {
                Console.WriteLine(e.Message);
                Console.ReadKey();
                return false;
            }

            return true;
        }

        static bool CanImport(Google.Apis.Drive.v3.Data.File file, ContentElement ele, ContentType type)
        {
            if (DriveHelper.IsSpreadsheet(file))
            {
                // Target element isn't needed for importing spreadsheets
                return (file != null && type != null);
            }
            else if (DriveHelper.IsAsset(file))
            {
                // Nothing needed for importing assets
                return true;
            }
            else
            {
                return (ele != null && file != null && type != null);
            }
        }

        public static void ShowError(string text)
        {
            WriteColoredText(text, ConsoleColor.Red);
        }

        public static void WriteColoredText(string text, ConsoleColor color)
        {
            Console.ForegroundColor = color;
            Console.WriteLine(text);
            Console.ForegroundColor = ConsoleColor.Gray;
        }

        static bool AskUpdate()
        {
            WriteColoredText("Would you like to update existing items if they exist, or create new items?", ConsoleColor.Green);
            Console.WriteLine("1. Update existing item");
            Console.WriteLine("2. Create new item");
            var input = Console.ReadLine();
            if (int.TryParse(input, out int num))
            {
                return (num == 1);
            }
            else
            {
                ShowError("Unrecognized number. Please press any key to try again.");
                Console.ReadKey();
                return AskUpdate();
            }
        }
        

    }

}
