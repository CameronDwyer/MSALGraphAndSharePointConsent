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
            const string clientId = "f70003ed-5a99-4600-a0d3-531950e2c805"; // OnePlace Solution Desktop Suite - Dev R85

            Console.WriteLine("MSAL consolidated Graph & SharePoint consent test app");
            Console.WriteLine("You should be prompted to login to M365, if you have no existing consent for this\napp you should be presented with scopes for both Graph and SharePoint.\n\n");

            publicClientApp = PublicClientApplicationBuilder.Create(clientId)
                .WithRedirectUri("http://localhost")
                .Build();

            var accounts = await publicClientApp.GetAccountsAsync();

            string[] graphScopes = new string[] { "https://graph.microsoft.com/user.read", "https://graph.microsoft.com/sites.readwrite.all" };
            string[] sharepointScopesForConsent = new string[] { "https://microsoft.sharepoint-df.com/allsites.manage" };

            // Make first call to get Graph access token (and get consent for all scopes needed Graph + SharePoint)
            AuthenticationResult graphAuthResult = await AuthenticateWithAzureADAsync(graphScopes, sharepointScopesForConsent, accounts.FirstOrDefault());

            // Make call to discover the root SharePoint site url from Graph
            string sharePointRootWebUrl = ((dynamic)JsonConvert.DeserializeObject(
                   await MakeGraphApiCall("https://graph.microsoft.com/v1.0/sites/root?$select=webUrl", graphAuthResult.AccessToken)
                   )).webUrl;

            Console.WriteLine($"[Graph API response]\nSharePoint Root Site WebUrl is: {sharePointRootWebUrl}\n");

            // Use MSAL to get SharePoint token
            // notice we should get no second prompt for consent we did it all without knowing the users SharePoint url
            accounts = await publicClientApp.GetAccountsAsync();
            string[] sharepointScopes = new string[] { $"{sharePointRootWebUrl}/allsites.manage" };
            AuthenticationResult sharePointAuthResult = await AuthenticateWithAzureADAsync(sharepointScopes, graphScopes, accounts.FirstOrDefault());          

            // Now make call to SharePoint REST API proving the token works
            string sharepointResponse = await MakeSharePointRestApiCall($"{sharePointRootWebUrl}/_api/web", sharePointAuthResult.AccessToken);
            Console.WriteLine($"[SharePoint REST API response] (first 500 chars):\n{sharepointResponse.Substring(0,500)}");

            Console.WriteLine("If you got here then we managed to get SharePoint consent without knowing your SharePoint root web URL!");
            Console.WriteLine("\nPress any key to finish");
            Console.ReadKey();
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
