using Google.Apis.Auth.OAuth2;
using Google.Apis.Drive.v3;
using Google.Apis.Services;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;

namespace DriveImportCore
{
    class DriveHelper
    {
        DriveService driveService;

        const string MimeTypeGoogleDoc = "application/vnd.google-apps.document",
        MimeTypeOpenOfficeSheet = "application/x-vnd.oasis.opendocument.spreadsheet",
        MimeTypeGoogleSheet = "application/vnd.google-apps.spreadsheet",
        MimeTypeExcelSheet = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        MimeTypeGooglePhoto = "application/vnd.google-apps.photo";

        /// <summary>
        /// MimeTypes which are automatically considered assets
        /// </summary>
        static string[] MimeTypeSupportedAssets = {
            "image/jpeg",
            "image/gif",
            "image/png",
            "image/svg+xml"
        };

        /// <summary>
        /// Extensions which require special handling when downloading from Drive
        /// </summary>
        string[] SpecialCases =
       {
            "txt"
        };
        string StandardFields = "files(id, name, sharedWithMeTime, ownedByMe, mimeType, fileExtension)";


        public DriveHelper(UserCredential credentials, string appName)
        {
            driveService = new DriveService(new BaseClientService.Initializer()
            {
                HttpClientInitializer = credentials,
                ApplicationName = appName,
            });
        }

        public List<Google.Apis.Drive.v3.Data.File> GetFilesInDirectory(string directory)
        {
            string dirID = null;
            FilesResource.ListRequest idRequest = GetListRequest(
                $"name = '{directory}' and trashed = false and mimeType = 'application/vnd.google-apps.folder'",
                StandardFields);

            IList<Google.Apis.Drive.v3.Data.File> dir = idRequest.Execute().Files;
            if (dir != null && dir.Count > 0)
            {
                dirID = dir[0].Id;
                // Obtain file list
                FilesResource.ListRequest dirList = GetListRequest(
                    $"'{dirID}' in parents and trashed = false and mimeType != 'application/vnd.google-apps.folder'",
                    StandardFields);

                var list = new List<Google.Apis.Drive.v3.Data.File>();
                IList<Google.Apis.Drive.v3.Data.File> dirFiles = dirList.Execute().Files;
                dirFiles = FilterFiles(dirFiles);
                if (dirFiles != null && dirFiles.Count > 0)
                {
                    foreach (var file in dirFiles)
                    {
                        list.Add(file);
                    }
                }

                return list;
            }
            else
            {
                throw new Exception($"Directory '{directory}' not found.");
            }
        }

        public Google.Apis.Drive.v3.Data.File GetSingleFile(string filename)
        {
            Google.Apis.Drive.v3.Data.File folder = null;
            if (filename.Contains("/"))
            {
                // Filename is a path (e.g. /folder1/folder2/myfile.txt)
                string[] path = filename.Split('/');
                // Get second-to-last part of path, aka the file's direct parent folder
                int index = path.Length - 2;
                if (index < 0)
                {
                    throw new Exception($"Invalid path {filename}.");
                }

                // Change filename to the last part of the path
                filename = path[path.Length - 1];
                var request = GetListRequest($"trashed = false and mimeType = 'application/vnd.google-apps.folder' and name = '{path[index]}'", StandardFields);
                var list = request.Execute().Files;
                folder = list[0];
            }

            string query = $"trashed = false and mimeType != 'application/vnd.google-apps.folder' and name = '{filename}'";
            if (folder != null) query += $" and '{folder.Id}' in parents";

            FilesResource.ListRequest listRequest = GetListRequest(query, StandardFields);
            var files = listRequest.Execute().Files;
            files = FilterFiles(files);
            if (files != null && files.Count > 0)
            {
                return files[0];
            }
            else
            {
                throw new Exception($"Source file '{filename}' not found.");
            }
        }

        

        /// <summary>
        /// List all available files on the Google Drive
        /// </summary>
        /// <returns></returns>
        public IList<Google.Apis.Drive.v3.Data.File> ListFiles()
        {
            FilesResource.ListRequest listRequest = GetListRequest(
                "trashed = false and mimeType != 'application/vnd.google-apps.folder'",
                StandardFields);
            IList<Google.Apis.Drive.v3.Data.File> files = listRequest.Execute().Files;
            return FilterFiles(files);
        }

        public MemoryStream DownloadFile(Google.Apis.Drive.v3.Data.File file)
        {
            Console.WriteLine("Downloading file '{0}'", file.Name);
            string mimeType = null;
            switch (file.MimeType)
            {
                case MimeTypeExcelSheet:
                    mimeType = MimeTypeExcelSheet;
                    break;
                case MimeTypeGoogleSheet:
                    mimeType = "text/csv";
                    break;
                case MimeTypeGoogleDoc:
                case "text/plain":
                    mimeType = "text/plain";
                    break;
            }

            if (mimeType == null && IsAsset(file))
            {
                mimeType = file.MimeType;
            }

            MemoryStream stream = null;
            if (mimeType == null)
            {
                // Unsupported extension
                Program.ShowError($"The MimeType {file.MimeType} is currently unsupported.");
                Console.ReadKey();
            }
            else
            {
                stream = new MemoryStream();
                if (IsAsset(file) || SpecialCases.Contains(file.FileExtension))
                {
                    var request = GetFileRequest(file);
                    request.Download(stream);
                }
                else
                {
                    var request = GetExportRequest(file, mimeType);
                    request.Download(stream);
                }
            }
            return stream;
        }

        public static bool IsSpreadsheet(Google.Apis.Drive.v3.Data.File file)
        {
            return (file.MimeType == MimeTypeGoogleSheet || file.MimeType == MimeTypeExcelSheet);
        }

        /// <summary>
        /// Returns true if the file should be uploaded to Kentico Kontent as an asset instead of content
        /// </summary>
        /// <param name="file"></param>
        /// <returns></returns>
        public static bool IsAsset(Google.Apis.Drive.v3.Data.File file)
        {
            return MimeTypeSupportedAssets.Contains(file.MimeType);
        }

        /// <summary>
        /// Removes files from the list that are "Shared with me" and not owned by user
        /// </summary>
        /// <param name="files"></param>
        /// <returns></returns>
        public List<Google.Apis.Drive.v3.Data.File> FilterFiles(IList<Google.Apis.Drive.v3.Data.File> files)
        {
            List<Google.Apis.Drive.v3.Data.File> result = new List<Google.Apis.Drive.v3.Data.File>();
            foreach (var f in files)
            {
                if (f.SharedWithMeTime != null) continue;
                else if (f.OwnedByMe == false) continue;
                else result.Add(f);
            }
            return result;
        }

        public FilesResource.GetRequest GetFileRequest(Google.Apis.Drive.v3.Data.File file)
        {
            return driveService.Files.Get(file.Id);
        }

        public FilesResource.ExportRequest GetExportRequest(Google.Apis.Drive.v3.Data.File file, string mimeType)
        {
            return driveService.Files.Export(file.Id, mimeType);
        }

        public FilesResource.ListRequest GetListRequest(string q, string fields)
        {
            var request = driveService.Files.List();
            request.Q = q;
            request.Fields = fields;
            return request;
        }
    }
}
