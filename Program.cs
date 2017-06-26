using Microsoft.IdentityModel.Clients.ActiveDirectory;
using Newtonsoft.Json;
using Newtonsoft.Json.Linq;
using System;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;
using System.Net.Http;
using System.Net.Http.Headers;
using System.Text;
using System.Threading.Tasks;

namespace SkypeUCWASendIM
{
    //https://ucwa.skype.com/documentation/keytasks-createapplication
    //https://code.msdn.microsoft.com/vstudio/Create-a-UCWA-Windows-2c48d3f9
    //https://www.matthewproctor.com/Send-An-IM-With-UCWA-Sending-The-IM/
    //https://blogs.msdn.microsoft.com/tsmatsuz/2017/04/13/skype-for-business-trusted-application-api-tutorial/
    //https://code.msdn.microsoft.com/vstudio/Create-a-UCWA-Windows-2c48d3f9    
    //https://code.msdn.microsoft.com/Lync-2013-Open-an-event-d4b6cb62
    //https://blog.missiaen.com/?cat=28
    class Program
    {
        private static string tenant = "xxx";
        private static string clientId = "xxx";
        private static string sfboResourceAppId = "00000004-0000-0ff1-ce00-000000000000";
        private static string aadInstance = "https://login.microsoftonline.com/{0}";
        private static string redirectUri = "https://3b14be72.ngrok.io/callback";
        private static AuthenticationContext authenticationContext = null;
        private static string hardcodedUsername = "xxx";
        private static string hardcodedPassword = "p@ssw0rd";
        private static string ucwaApplicationsUri = "";
        private static string ucwaApplicationsHost = "";
        private static string createUcwaAppsResults = "";
        private static AuthenticationResult ucwaAuthenticationResult = null;
        private static string Accesstoken = string.Empty;
        static void Main(string[] args)
        {
            var dt = new DateTime(1498464308079);
            authenticationContext = new AuthenticationContext
                 (String.Format(CultureInfo.InvariantCulture, aadInstance, tenant));
            authenticationContext.TokenCache.Clear();
            AuthenticationResult testCredentials = null;
            UserCredential uc = null;
            uc = new UserPasswordCredential(hardcodedUsername, hardcodedPassword);
            testCredentials = GetAzureAdToken(authenticationContext, sfboResourceAppId, clientId, redirectUri, uc);
            Accesstoken = testCredentials.AccessToken;
            ucwaApplicationsUri = GetUcwaRootUri(authenticationContext, sfboResourceAppId, clientId, redirectUri, uc);

            ucwaApplicationsHost = ReduceUriToProtoAndHost(ucwaApplicationsUri);
            ucwaAuthenticationResult = GetAzureAdToken(authenticationContext, ucwaApplicationsHost, clientId, redirectUri, uc);
            Accesstoken = ucwaAuthenticationResult.AccessToken;

            UcwaMyAppsObject ucwaMyAppsObject = new UcwaMyAppsObject()
            {
                UserAgent = "Notification_App",
                EndpointId = Guid.NewGuid().ToString(),
                Culture = "en-US"
            };

            Console.WriteLine("Making request to ucwaApplicationsUri " + ucwaApplicationsUri);
            createUcwaAppsResults = CreateUcwaApps(ucwaAuthenticationResult, ucwaApplicationsUri, ucwaMyAppsObject);

            SendIMToUser(); // send im to user
        }

        public static AuthenticationResult GetAzureAdToken(AuthenticationContext authContext, String resourceHostUri,
           string clientId, string redirectUri, UserCredential uc)
        {

            AuthenticationResult authenticationResult = null;

            Console.WriteLine("Performing GetAzureAdToken");
            try
            {
                Console.WriteLine("Passed resource host URI is " + resourceHostUri);
                if (resourceHostUri.StartsWith("http"))
                {
                    resourceHostUri = ReduceUriToProtoAndHost(resourceHostUri);
                    Console.WriteLine("Normalized the resourceHostUri to just the protocol and hostname " + resourceHostUri);
                }

                // check if there's a user credential - i.e. a username and password

                if (uc != null)
                {
                    // ClientCredential cc = new ClientCredential("ed95a52c-9950-4dbb-bfef-c33cb78a5ba4",
                    //                        "anVzNjaUKtIsXGhrKnO+mpffrsoFUzkOb51F1/5Z37E=");
                    authenticationResult = authContext.AcquireTokenAsync(resourceHostUri, clientId, uc).Result;
                    //authContext.AcquireTokenAsync(resourceHostUri,cc).Result;
                    //

                }
                else
                {
                    PlatformParameters platformParams = new PlatformParameters(PromptBehavior.Auto);
                    authenticationResult = authContext.AcquireTokenAsync(resourceHostUri, clientId, new Uri(redirectUri), platformParams).Result;
                }

                //Console.WriteLine("Bearer token from Azure AD is " + authenticationResult.AccessToken);
            }
            catch (Exception ex)
            {
                Console.ForegroundColor = ConsoleColor.Red;
                Console.WriteLine("An unexpected error occurred.");
                string message = ex.Message;
                if (ex.InnerException != null)
                {
                    message += Environment.NewLine + "Inner Exception : " + ex.InnerException.Message;
                }
                Console.WriteLine("Message: {0}", message);
                Console.ForegroundColor = ConsoleColor.White;

            }

            return authenticationResult;
        }

        public static String ReduceUriToProtoAndHost(string longUri)
        {
            string reduceUriToProtoAndHost = String.Empty;
            reduceUriToProtoAndHost = new Uri(longUri).Scheme + "://" + new Uri(longUri).Host;
            return reduceUriToProtoAndHost;
        }

        public static HttpClient SharedHttpClient = new HttpClient();
        private static string ucwaAutoDiscoveryUri = "https://webdir.online.lync.com/autodiscover/autodiscoverservice.svc/root";

        public static string GetUcwaRootUri(AuthenticationContext authenticationContext, String sfboResourceAppId,
                string clientId, string redirectUri, UserCredential uc)
        {
            Console.WriteLine("Now we'll call UCWA Autodiscovery to get the root/oauth/user URI");
            var ucwaAutoDiscoveryUserRootUri = DoUcwaAutoDiscovery(SharedHttpClient, authenticationContext, sfboResourceAppId, clientId, redirectUri, uc);

            Console.WriteLine("Now we'll get the UCWA Applications URI for the user");
            var ucwaRootUri = GetUcwaUserResourceUri(SharedHttpClient, authenticationContext, ucwaAutoDiscoveryUserRootUri, clientId, redirectUri, uc);

            return ucwaRootUri;
        }
        private static string DoUcwaAutoDiscovery(HttpClient httpClient, AuthenticationContext authenticationContext, String sfboResourceAppId, string clientId, string redirectUri, UserCredential uc)
        {
            //AuthenticationResult authenticationResult = null;
            //authenticationResult = GetAzureAdToken(authenticationContext, sfboResourceAppId, clientId, redirectUri, uc);
            //https://webdir.online.lync.com/autodiscover/autodiscoverservice.svc/root
            string ucwaAutoDiscoveryUserRootUri = string.Empty;

            //Console.WriteLine("Using this access token " + result.AccessToken);
            httpClient.DefaultRequestHeaders.Clear();
            httpClient.DefaultRequestHeaders.Authorization = new AuthenticationHeaderValue("Bearer", Accesstoken);
            var httpResponseMessage = httpClient.GetAsync(ucwaAutoDiscoveryUri).Result;
            if (httpResponseMessage.IsSuccessStatusCode)
            {
                Console.WriteLine("Called " + ucwaAutoDiscoveryUri);
                var resultString = httpResponseMessage.Content.ReadAsStringAsync().Result;
                Console.WriteLine("DoUcwaDiscovery URI " + resultString);
                dynamic resultObject = JObject.Parse(resultString);
                ucwaAutoDiscoveryUserRootUri = resultObject._links.user.href;
                Console.WriteLine("DoUcwaDiscovery Root URI is " + ucwaAutoDiscoveryUserRootUri);
            }
            return ucwaAutoDiscoveryUserRootUri;
        }

        private static string GetUcwaUserResourceUri(HttpClient httpClient, AuthenticationContext authenticationContext, String ucwaUserDiscoveryUri, string clientId,
            string redirectUri, UserCredential uc)
        {
            //https://webdir1a.online.lync.com/Autodiscover/AutodiscoverService.svc/root/oauth/user
            AuthenticationResult authenticationResult = null;
            authenticationResult = GetAzureAdToken(authenticationContext, ucwaUserDiscoveryUri, clientId, redirectUri, uc);

            string ucwaUserResourceUri = String.Empty;

            httpClient.DefaultRequestHeaders.Clear();
            httpClient.DefaultRequestHeaders.Authorization = new AuthenticationHeaderValue("Bearer", authenticationResult.AccessToken);
            var httpResponseMessage = httpClient.GetAsync(ucwaUserDiscoveryUri).Result;

            Console.WriteLine("Called " + ucwaUserDiscoveryUri);
            var resultString = httpResponseMessage.Content.ReadAsStringAsync().Result;
            Console.WriteLine("GetUcwaUserResourceUri Body " + resultString);
            dynamic resultObject = JObject.Parse(resultString);
            string resourceRedirectUri = "";
            try
            {
                resourceRedirectUri = resultObject._links.redirect.href;
            }
            catch
            {
                Console.WriteLine("No re-direct");
            }
            if (resourceRedirectUri != "")
            {
                Console.WriteLine("GetUcwaUserResourceUri redirect is " + resourceRedirectUri);
                resourceRedirectUri += "/oauth/user";  // for some reason, the redirectUri doesn't include /oauth/user
                Console.WriteLine("Modifying GetUcwaUserResourceUri to be correct " + resourceRedirectUri);
                // recursion is your friend
                ucwaUserResourceUri = GetUcwaUserResourceUri(httpClient, authenticationContext, resourceRedirectUri, clientId, redirectUri, uc);
            }
            else  // if there's no redirect then the applications URI is there for us to grab
            {
                ucwaUserResourceUri = resultObject._links.applications.href;
            }
            Console.WriteLine("GetUcwaUserResourceUri is " + ucwaUserResourceUri);
            return ucwaUserResourceUri;
        }

        public class UcwaMyAppsObject
        {
            public string UserAgent { get; set; }
            public string EndpointId { get; set; }
            public string Culture { get; set; }
        }

        public static string CreateUcwaApps(AuthenticationResult ucwaAuthenticationResult, string ucwaApplicationsRootUri,
            UcwaMyAppsObject ucwaAppsObject)
        {
            string createUcwaAppsResults = string.Empty;
            SharedHttpClient.DefaultRequestHeaders.Clear();
            SharedHttpClient.DefaultRequestHeaders.Authorization = new AuthenticationHeaderValue("Bearer", ucwaAuthenticationResult.AccessToken);
            SharedHttpClient.DefaultRequestHeaders.Accept.Add(new MediaTypeWithQualityHeaderValue("application/json"));
            var createUcwaPostData = JsonConvert.SerializeObject(ucwaAppsObject);
            Console.WriteLine("CreateUcwaApps POST data is " + createUcwaPostData);
            var httpResponseMessage =
                SharedHttpClient.PostAsync(ucwaApplicationsRootUri, new StringContent(createUcwaPostData, Encoding.UTF8,
                "application/json")).Result;
            Console.WriteLine("CreateUcwaApps response is " + httpResponseMessage.Content.ReadAsStringAsync().Result);
            if (httpResponseMessage.IsSuccessStatusCode)
            {
                createUcwaAppsResults = httpResponseMessage.Content.ReadAsStringAsync().Result;
            }

            return createUcwaAppsResults;
        }

        static void SendIMToUser()
        {
            dynamic createUcwaAppsResultsObject = JObject.Parse(createUcwaAppsResults);
            string sendIMUri = ucwaApplicationsHost +
                   createUcwaAppsResultsObject._embedded.communication._links.startMessaging.href;

            SharedHttpClient.DefaultRequestHeaders.Clear();
            SharedHttpClient.DefaultRequestHeaders.Authorization = new AuthenticationHeaderValue("Bearer", ucwaAuthenticationResult.AccessToken);
            SharedHttpClient.DefaultRequestHeaders.Accept.Add(new MediaTypeWithQualityHeaderValue("application/json"));
                var PostData = JsonConvert.SerializeObject( new SendIM
                (){
                                                importance= "Normal",
                                                sessionContext=Guid.NewGuid().ToString(),
                                                subject="Task Sample",
                                                telemetryId=null,
                                                to= "sip:user1@tesla329.onmicrosoft.com",
                                                operationId= Guid.NewGuid().ToString()
                });

            var httpResponseMessage =
                SharedHttpClient.PostAsync(sendIMUri, new StringContent(PostData, Encoding.UTF8,
                "application/json")).Result;
            Console.WriteLine("Send IM response is " + httpResponseMessage.Content.ReadAsStringAsync().Result);

            SharedHttpClient.DefaultRequestHeaders.Clear();
            SharedHttpClient.DefaultRequestHeaders.Authorization = new AuthenticationHeaderValue("Bearer", ucwaAuthenticationResult.AccessToken);
            SharedHttpClient.DefaultRequestHeaders.Accept.Add(new MediaTypeWithQualityHeaderValue("application/json"));
            var getEvent = ucwaApplicationsHost +
                   createUcwaAppsResultsObject._links.events.href;
            httpResponseMessage =
                SharedHttpClient.GetAsync(getEvent).Result;
            dynamic nextEventReesponse=null;
            string nextEvent=null;
            if (httpResponseMessage.IsSuccessStatusCode)
            {
                var result = httpResponseMessage.Content.ReadAsStringAsync().Result;
                Console.WriteLine("IM response is " + result);
                nextEventReesponse=JObject.Parse(result);
                nextEvent = nextEventReesponse._links.next.href;
            }

            SharedHttpClient.DefaultRequestHeaders.Clear();
            SharedHttpClient.DefaultRequestHeaders.Authorization = new AuthenticationHeaderValue("Bearer", ucwaAuthenticationResult.AccessToken);
            SharedHttpClient.DefaultRequestHeaders.Accept.Add(new MediaTypeWithQualityHeaderValue("application/json"));
           
            httpResponseMessage =
                SharedHttpClient.GetAsync(ucwaApplicationsHost + nextEvent).Result;
            
            if (httpResponseMessage.IsSuccessStatusCode)
            {
                var result = httpResponseMessage.Content.ReadAsStringAsync().Result;
                Console.WriteLine("IM response is " + result);
                nextEventReesponse = JObject.Parse(result);
                nextEvent = nextEventReesponse._links.next.href;
            }

            SharedHttpClient.DefaultRequestHeaders.Clear();
            SharedHttpClient.DefaultRequestHeaders.Authorization = new AuthenticationHeaderValue("Bearer", ucwaAuthenticationResult.AccessToken);
            SharedHttpClient.DefaultRequestHeaders.Accept.Add(new MediaTypeWithQualityHeaderValue("application/json"));
             httpResponseMessage =
                 SharedHttpClient.PostAsync(ucwaApplicationsHost+ nextEventReesponse.sender[1].events[3]._embedded.messaging._links.sendMessage.href , new StringContent("PLease Approve", Encoding.UTF8,
                 "text/plain")).Result;

            Console.WriteLine("Send IM response is " + httpResponseMessage.Content.ReadAsStringAsync().Result);
            //text/html
            //https://www.matthewproctor.com/Send-An-IM-With-UCWA-Sending-the-IM/
        }

        public class SendIM
        {
            public string importance { get; set; }
            public string sessionContext { get; set; }
            public string subject { get; set; }
            public string telemetryId { get; set; }
            public string to { get; set; }
            public string operationId { get; set; }
            
        }
    }
}
