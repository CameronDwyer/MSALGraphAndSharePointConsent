using Microsoft.Identity.Client;
using Newtonsoft.Json;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Net.Http;
using System.Net.Http.Headers;
using System.Threading.Tasks;

namespace MSALConsolidatedGraphAndSharePointConsent
{
    class Program
    {
        static IPublicClientApplication publicClientApp;

        static async Task Main(string[] args)
        {
            const string clientId = "7b46c75f-8bcf-439e-a0fd-9afe2128dd5a"; // OnePlace Solution Desktop Suite - Dev R85

            Console.WriteLine("MSAL consolidated Graph & SharePoint consent test app");
            Console.WriteLine("You should be prompted to login to M365, if you have no existing consent for this\napp you should be presented with scopes for both Graph and SharePoint.\n\n");

            publicClientApp = PublicClientApplicationBuilder.Create(clientId)
                .WithRedirectUri("http://localhost")
                .Build();

            var accounts = await publicClientApp.GetAccountsAsync();

            string[] graphScopes = new string[] { "https://graph.microsoft.com/user.read", "https://graph.microsoft.com/sites.readwrite.all" };
            string[] sharepointScopes = new string[] { "https://microsoft.sharepoint-df.com/allsites.manage" };

            // Make first call to get Graph access token (and get consent for all scopes needed Graph + SharePoint)
            AuthenticationResult authResultGraph = await AuthenticateWithAzureADAsync(graphScopes, sharepointScopes, accounts.FirstOrDefault());

            // Make call to discover the root SharePoint site url from Graph
            string sharePointRootWebUrl = ((dynamic)JsonConvert.DeserializeObject(await MakeGraphApiCall("https://graph.microsoft.com/v1.0/sites/root?$select=webUrl", authResultGraph.AccessToken))).webUrl;
            Console.Write($"SharePoint Root Site WebUrl is: {sharePointRootWebUrl}");


        }

        private static async Task<AuthenticationResult> AuthenticateWithAzureADAsync(IEnumerable<string> scopes, IEnumerable<string> extraResourceScopesToConsent, IAccount account)
        {
            if (account == null)
            {
                return await publicClientApp
                    .AcquireTokenInteractive(scopes)
                    .WithExtraScopesToConsent(extraResourceScopesToConsent)
                    .WithAuthority($"https://login.microsoftonline.com/organizations")
                    .ExecuteAsync();
            }
            else
            {
                try
                {
                    return await publicClientApp
                        .AcquireTokenSilent(scopes, account)
                        .WithAuthority($"https://login.microsoftonline.com/organizations")
                        .ExecuteAsync();
                }
                catch (MsalUiRequiredException ex)
                {
                    return await publicClientApp
                        .AcquireTokenInteractive(scopes)
                        .WithExtraScopesToConsent(extraResourceScopesToConsent)
                        .WithAccount(account)
                        .ExecuteAsync();
                }
            }
        }

        private static async Task<string> MakeGraphApiCall(string uri, string accessToken)
        {
            HttpClient graphHttpClient = new HttpClient();
            graphHttpClient.DefaultRequestHeaders.Authorization = new AuthenticationHeaderValue("Bearer", accessToken);
            return await graphHttpClient.GetStringAsync(uri);

        }

        private static async Task<string> MakeSharePointRestApiCall(string uri, string accessToken)
        {
            HttpClient sharepointHttpClient = new HttpClient();
            sharepointHttpClient.DefaultRequestHeaders.Authorization = new AuthenticationHeaderValue("Bearer", accessToken);
            return await sharepointHttpClient.GetStringAsync(uri);
        }
    }
}
