using Google.Apis.Auth.OAuth2;
using Google.Apis.Drive.v3;
using Google.Apis.Services;

namespace BookMarkApp.Models
{
    public class DriveV3Snippets
    {
        public static string DriveUploadBasic(string serviceAccountKeyFilePath, string folderId, string filePath)
        {
            try
            {
                GoogleCredential credential;

                using (var stream = new FileStream(serviceAccountKeyFilePath, FileMode.Open, FileAccess.Read))
                {
                    credential = GoogleCredential.FromStream(stream)
                        .CreateScoped(DriveService.Scope.DriveFile);
                }

                var driveService = new DriveService(new BaseClientService.Initializer
                {
                    HttpClientInitializer = credential,
                    ApplicationName = "bookmark",
                });

                var fileMetadata = new Google.Apis.Drive.v3.Data.File
                {
                    Name = Path.GetFileName(filePath),
                    Parents = string.IsNullOrEmpty(folderId) ? null : new[] { folderId },
                };

                FilesResource.CreateMediaUpload request;

                using (var stream = new FileStream(filePath, FileMode.Open))
                {
                    request = driveService.Files.Create(fileMetadata, stream, "application/octet-stream");
                    request.Upload();
                }

                var file = request.ResponseBody;
               return $"File uploaded: {file.Name} (ID: {file.Id})";
            }
            catch (Exception ex)
            {
                return $"Error: {ex.Message}";
            }
        }
    }
}

