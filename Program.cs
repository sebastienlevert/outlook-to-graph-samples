using System;
using System.Net.Http;
using System.Threading.Tasks;
using Newtonsoft.Json;

namespace OutlookToGraphSamples
{
    class Program
    {
        static DeviceCodeAuthProvider authProvider = null;
        static Guid clientId = new Guid("00000000-0000-0000-0000-000000000000");
        static string[] outlookScopes = new string[] { "https://outlook.office.com/User.Read", "https://outlook.office.com/Mail.Read" };
        static string[] graphScopes = new string[] { "https://graph.microsoft.com/User.Read", "https://graph.microsoft.com/Mail.Read" };
        static async Task Main(string[] args)
        {
            Console.WriteLine("================================================");
            Console.WriteLine("Welcome to the Outlook to Microsoft Graph Sample");
            Console.WriteLine("================================================");

            #region Initialization
            var argClientId = args.Length > 0 ? args[0] : clientId.ToString();
            try {
                clientId = Guid.Parse(argClientId);
                authProvider = new DeviceCodeAuthProvider(clientId.ToString(), outlookScopes);
            } catch (FormatException) {
                Console.WriteLine($"Bad clientId format: {argClientId}");
                Environment.Exit(-1);
            }
            
            Console.WriteLine("------------------------------------------------");
            Console.WriteLine($"Using the following Client Id: {clientId}");
            Console.WriteLine("------------------------------------------------");
            #endregion
            
            Console.WriteLine("------------------------------------------------");
            Console.WriteLine("Outlook REST APIs : Getting the latest messages ");
            Console.WriteLine("------------------------------------------------");
            
            Console.WriteLine($"Authenticating using following scopes: {String.Join(", ", outlookScopes)}");
            string outlookAccessToken = GetAccessToken(outlookScopes);
            var outlookMessages = await GetMessagesViaOutlookRestAPIs(outlookAccessToken);
            Console.WriteLine($"Message Count: {outlookMessages.Count}");
                        
            Console.WriteLine("------------------------------------------------");
            Console.WriteLine("Microsoft Graph : Getting the latest messages ");
            Console.WriteLine("------------------------------------------------");
            Console.WriteLine($"Authenticating using following scopes: {String.Join(", ", graphScopes)}");
            string graphAccessToken = GetAccessToken(graphScopes);
            var graphMessages = await GetMessagesViaGraph(graphAccessToken);
            Console.WriteLine($"Message Count: {graphMessages.Count}");
        }

        static async Task<dynamic> GetMessagesViaOutlookRestAPIs(string accessToken) {
            var response = await MakeApiCall("GET", accessToken, "https://outlook.office.com/api/v2.0/me/messages");
            string data = await response.Content.ReadAsStringAsync();
            dynamic jsonData = JsonConvert.DeserializeObject<dynamic>(data);
            return jsonData.value;        
        }

          static async Task<dynamic> GetMessagesViaGraph(string accessToken) {
            var response = await MakeApiCall("GET", accessToken, "https://graph.microsoft.com/v1.0/me/messages");
            string data = await response.Content.ReadAsStringAsync();
            dynamic jsonData = JsonConvert.DeserializeObject<dynamic>(data);
            return jsonData.value;            
        }
        
        static string GetAccessToken(string[] scopes) {
            if(authProvider == null) {
                authProvider = new DeviceCodeAuthProvider(clientId.ToString(), scopes);
            }   

            var accessToken = authProvider.GetAccessToken(scopes).Result;

            return accessToken;
        }

        public static async Task<HttpResponseMessage> MakeApiCall(string method, string token, string apiUrl)
        {
            using (var httpClient = new HttpClient())
            {
                var request = new HttpRequestMessage(new HttpMethod(method), apiUrl);

                // Headers
                request.Headers.Authorization = new System.Net.Http.Headers.AuthenticationHeaderValue("Bearer", token);
                
                var apiResult = await httpClient.SendAsync(request);
                return apiResult;
            }
        }
    }
}
