using Google.Apis.Auth.OAuth2;
using Google.Apis.Services;
using Google.Apis.Sheets.v4;
using System;
using System.Collections.Generic;
using System.IO;
using System.Text;

namespace NightlyRouteToSlack 
{
    class GoogleService
    {
        private readonly string _googleSecretJsonFilePath;
        private readonly string _applicationName;
        private readonly string[] _scopes;

        public GoogleService(string googleSecretJsonFilePath, string applicationName, string[] scopes)
        {
            _googleSecretJsonFilePath = googleSecretJsonFilePath;
            _applicationName = applicationName;
            _scopes = scopes;
        }

        public GoogleCredential GetGoogleCredential()
        {
            GoogleCredential credential;
            using (var stream =
                new FileStream(_googleSecretJsonFilePath, FileMode.Open, FileAccess.Read))
            {

                credential = GoogleCredential.FromStream(stream).CreateScoped(_scopes);
            }
            return credential;
        }

        public SheetsService GetSheetsService()
        {
            var credential = GetGoogleCredential();
            var sheetsService = new SheetsService(new BaseClientService.Initializer()
            {
                HttpClientInitializer = credential,
                ApplicationName = _applicationName,
            });
            return sheetsService;
        }
    }
}
