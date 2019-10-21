using KenticoCloud.ContentManagement;
using KenticoCloud.ContentManagement.Exceptions;
using KenticoCloud.ContentManagement.Models.Assets;
using KenticoCloud.ContentManagement.Models.Items;
using Kentico.Kontent.Delivery;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text.RegularExpressions;
using System.Threading.Tasks;

namespace DriveImportCore
{
    class KontentHelper
    {
        const string HTML_TAG_PATTERN = "<.*?>";
        const string SPREADSHEET_CODENAME_HEADER = "Name";
        static IDeliveryClient clientDelivery;
        static ContentManagementClient clientCM;

        static ImportOptions importOptions;

        public static void Init(IDeliveryClient delclient, ContentManagementClient cmclient)
        {
            clientDelivery = delclient;
            clientCM = cmclient;
        }

        /// <summary>
        /// Processes MemoryStream and upserts data. If the file is a spreadsheet, processes the sheet into CSV and upserts individual rows. If the file is an asset, creates new asset.
        /// </summary>
        /// <param name="stream"></param>
        /// <param name="file"></param>
        /// <param name="element"></param>
        /// <param name="type"></param>
        /// <param name="update"></param>
        public static void BeginImport(MemoryStream stream, Google.Apis.Drive.v3.Data.File file, ContentElement element, ContentType type, bool update)
        {
            string result = "";

            if (!DriveHelper.IsAsset(file))
            {
                // Read stream to string for content imports only
                stream.Seek(0, SeekOrigin.Begin);
                var sr = new StreamReader(stream);
                result = sr.ReadToEnd();
            }

            string importtype = DriveHelper.IsAsset(file) ? "Asset" : "Content";
            Program.WriteColoredText($"Download success, importing {importtype.ToLower()}..", ConsoleColor.Green);
                    
            if (DriveHelper.IsSpreadsheet(file))
            {
                // Import multiple items

                string[] rows = result.Split("\r\n");
                string[] headers = rows[0].Split(',');
                int codenamecolumn = Array.IndexOf(headers, SPREADSHEET_CODENAME_HEADER);
                for (var i = 1; i < rows.Length; i++)
                {
                    string[] rowdata = rows[i].Split(',');
                    if(rowdata.Length > headers.Length)
                    {
                        Program.ShowError($"Inconsistent data for row {i}. This could be because the row contains a comma, which is not currently supported. Skipping row.");
                        continue;
                    }
                    UpsertRowData(rowdata, headers, codenamecolumn, type, update);
                }
            }
            else if(DriveHelper.IsAsset(file)) {
                stream.Position = 0;
                UpsertAsset(file, update, stream);
            }
            else
            {
                // Import single item
                Dictionary<string, object> elements = null;
                if (element != null) {
                    elements = new Dictionary<string, object>()
                    {
                        {element.Codename, result}
                    };
                } 

                UpsertSingleItem(elements, file.Name, type, update);
            }
        }

        static string UpsertAsset(Google.Apis.Drive.v3.Data.File file, bool update, Stream stream)
        {
            string externalId = file.Name;
            string message = null;
            FileContentSource fcs = new FileContentSource(stream, externalId, file.MimeType);
            AssetModel result;
            var emptyDescriptions = new List<AssetDescription>
            {
                new AssetDescription { Description = "", Language = LanguageIdentifier.ById(new Guid("00000000-0000-0000-0000-000000000000")) }
            };

            try
            {
                var fileTask = clientCM.UploadFileAsync(fcs);
                FileReference fr = fileTask.GetAwaiter().GetResult();

                IEnumerable<AssetDescription> descriptions = new List<AssetDescription>();
                AssetUpsertModel model = new AssetUpsertModel
                {
                    FileReference = fr,
                    Descriptions = emptyDescriptions
                };
                var createTask = clientCM.CreateAssetAsync(model);
                result = createTask.GetAwaiter().GetResult();
            }
            catch(Exception ex)
            {
                message = ex.Message;
            }

            return message;
        }

        static void UpsertRowData(string[] rowdata, string[] headers, int codenamecolumn, ContentType type, bool update)
        {
            string codename = rowdata[codenamecolumn];
            Dictionary<string, object> elements = new Dictionary<string, object>();

            // Loop through row data and create an element with name equal to header and value equal to cell value
            for (var i = 0; i < rowdata.Length; i++)
            {
                if (i == codenamecolumn) continue;
                elements.Add(headers[i].ToLower(), rowdata[i]);
            }

            UpsertSingleItem(elements, codename, type, update);
        }

        static string StripHTML(string inputString)
        {
            return Regex.Replace(inputString, HTML_TAG_PATTERN, string.Empty);
        }

        static string Sanitize(string input)
        {
            string output = input.Replace("'", @"\'");
            return output;
        }

        private class ImportOptions
        {
            public string sourceFileName = null,
            targetElementName = null,
            contentTypeName = null,
            directory = null;
            public bool update = false;

            public ImportOptions(string[] args)
            {
                LoadImportOptions(args);
            }

            private void LoadImportOptions(string[] args)
            {
                for (int i = 0; i < args.Length; i++)
                {
                    if (args[i].StartsWith("-"))
                    {
                        switch (args[i])
                        {
                            // Name of Drive file to import
                            case "-source":
                            case "-s":
                                sourceFileName = (i + 1 <= args.Length - 1) ? args[i + 1] : null;
                                break;
                            // Name of folder to import contents from (overrides -source param)
                            case "-dir":
                            case "-d":
                                directory = (i + 1 <= args.Length - 1) ? args[i + 1] : null;
                                break;
                            // Name of content type to create/update
                            case "-type":
                            case "-t":
                                contentTypeName = (i + 1 <= args.Length - 1) ? args[i + 1] : null;
                                break;
                            // Name of element content will be placed
                            case "-element":
                            case "-e":
                                targetElementName = (i + 1 <= args.Length - 1) ? args[i + 1] : null;
                                break;
                            // If present, updateExisting
                            case "-update":
                            case "-u":
                                update = true;
                                break;
                        }
                    }
                }
            }
        }

        public static void AutoImport(string[] args, DriveHelper driveHelper)
        {
            ContentElement contentElement = null;
            ContentType contentType = null;
            
            importOptions = new ImportOptions(args);
            bool hasError = false;
            try
            {
                List<Google.Apis.Drive.v3.Data.File> filestoimport;
                if (importOptions.directory != null)
                {
                    filestoimport = driveHelper.GetFilesInDirectory(importOptions.directory);
                }
                else
                {
                    var file = driveHelper.GetSingleFile(importOptions.sourceFileName);
                    filestoimport = new List<Google.Apis.Drive.v3.Data.File>();
                    filestoimport.Add(file);
                }

                // Get content type
                var task = clientDelivery.GetTypeAsync(importOptions.contentTypeName);
                contentType = task.GetAwaiter().GetResult();

                // Obtain target element from content type
                contentElement = GetElementFromType(contentType, importOptions.targetElementName);
                
                //Import file(s)
                foreach (var file in filestoimport)
                {
                    var stream = driveHelper.DownloadFile(file);
                    if (stream != null) BeginImport(stream, file, contentElement, contentType, importOptions.update);
                }

            }
            catch (Exception e)
            {
                hasError = true;
                Program.ShowError(e.Message);
            }

            Console.WriteLine(hasError ? "Import not complete, please check error messages" : "Import complete");
            Console.ReadKey();
        }

        static ContentElement GetElementFromType(ContentType type, string targetelementname)
        {
            ContentElement element = null;

            var elements = FilterElements(type.Elements);
            foreach (var e in elements)
            {
                if (e.Key == targetelementname) element = e.Value;
            }
            if (element == null)
            {
                throw new Exception($"Element '{targetelementname}' not found. Ensure it exists in content type '{type}' and is either a Text or Rich Text element.");
            }
            else return element;
        }

        /// <summary>
        /// Returns only elements from the content type that are Text or Rich Text
        /// </summary>
        /// <param name="keys"></param>
        /// <returns></returns>
        public static List<KeyValuePair<string, ContentElement>> FilterElements(IReadOnlyDictionary<string, ContentElement> elements)
        {
            return elements.Where(e => e.Value.Type == "text" || e.Value.Type == "rich_text").ToList();
        }

        

        public static DeliveryTypeListingResponse ListContentTypes()
        {
            return clientDelivery.GetTypesAsync().Result;
        }

        static ContentItemModel CreateNewContentItem(string codename, string type)
        {
            ContentItemCreateModel newitem = new ContentItemCreateModel()
            {
                Name = codename,
                Type = ContentTypeIdentifier.ByCodename(type)
            };

            Task<ContentItemModel> createtask = clientCM.CreateContentItemAsync(newitem);
            return createtask.GetAwaiter().GetResult();
        }
        
        static ContentItem GetExistingContentItem(string codename, string contenttype)
        {
            var task = clientDelivery.GetItemsAsync(
                new EqualsFilter("system.type", contenttype),
                new EqualsFilter("system.name", codename)
            );
            var result = task.GetAwaiter().GetResult();
            if (result.Items.Count > 0)
            {
                return result.Items[0];
            }
            else
            {
                return null;
            }
        }

        /// <summary>
        /// </summary>
        /// <param name="elements">The Elements property of the content item to be updated</param>
        /// <param name="codename">If set, the item with the specified code name will be upserted</param>
        static void UpsertSingleItem(Dictionary<string, object> elements, string codename, ContentType type, bool update)
        {
            ContentItemModel contentItem = null;
            ContentItem existingItem = null;
            Guid itemid;

            // Try to get existing content item for updating
            if (update && existingItem == null) existingItem = GetExistingContentItem(codename, type.System.Codename);

            // If not updating, create content item first
            if (!update)
            {
                contentItem = CreateNewContentItem(codename, type.System.Codename);
                if (contentItem.Id != null)
                {
                    itemid = contentItem.Id;
                }
                else
                {
                    throw new Exception("Error creating new item.");
                }
            }
            else
            {
                // We are updating existing
                if (existingItem != null)
                {
                    itemid = new Guid(existingItem.System.Id);
                }
                else
                {
                    // Existing item wasn't found, create it
                    contentItem = CreateNewContentItem(codename, type.System.Codename);
                    if (contentItem.Id != null)
                    {
                        itemid = contentItem.Id;
                    }
                    else
                    {
                        throw new Exception("Error creating new item.");
                    }
                }
            }
            
            // After item is created (or skipped for updateExisting), upsert content
            try
            {
                // Get item variant to upsert
                ContentItemIdentifier itemIdentifier = ContentItemIdentifier.ById(itemid);
                LanguageIdentifier languageIdentifier = LanguageIdentifier.ById(new Guid("00000000-0000-0000-0000-000000000000"));
                ContentItemVariantIdentifier identifier = new ContentItemVariantIdentifier(itemIdentifier, languageIdentifier);

                elements = ValidateContentTypeFields(elements, type);

                // Set target element value
                ContentItemVariantUpsertModel model = new ContentItemVariantUpsertModel()
                {
                    Elements = elements
                };

                // Upsert item
                var upsertTask = clientCM.UpsertContentItemVariantAsync(identifier, model);
                var response = upsertTask.GetAwaiter().GetResult();
            }
            catch (ContentManagementException e)
            {
                if (e.Message.ToLower().Contains("cannot update published"))
                {
                     throw new Exception("This tool cannot currently update published content. If you wish to update a published item, you will first need to unpublish it within Kentico Kontent.");
                }
            }
        }

        /// <summary>
        /// For the targeted content type, ensures that the elements being updated exist and are sanitized
        /// </summary>
        static Dictionary<string, object> ValidateContentTypeFields(Dictionary<string, object> elements, ContentType type)
        {
            Dictionary<string, object> result = new Dictionary<string, object>();

            // Get target content type's elements
            var filteredelements = FilterElements(type.Elements);
            foreach (var ele in elements)
            {
                if (filteredelements.Any(f => f.Key == ele.Key))
                {
                    // Current element exists in target content type
                    // Strip HTML from value of element
                    string value = StripHTML(ele.Value.ToString());
                    var type_element = type.Elements[ele.Key];
                    if(type_element.Type == "rich_text")
                    {
                        // Surround rich text strings in <p> tag
                        value = "<p>" + value + "</p>";
                    }

                    result.Add(ele.Key, value);
                }

            }
            return result;
        }
    }
}
