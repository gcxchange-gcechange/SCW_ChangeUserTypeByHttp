using System.Linq;
using System.Net;
using System.Net.Http;
using System.Threading.Tasks;
using Microsoft.Azure.WebJobs;
using Microsoft.Azure.WebJobs.Extensions.Http;
using Microsoft.Azure.WebJobs.Host;
using System.Configuration;
using System;
using System.Text;
using System.Web;
using System.IO;
using System.Web.Script.Serialization;
using Microsoft.Graph;
using System.Net.Http.Headers;
using System.Collections.Generic;

namespace ChangeUserTypeByHttp
{
    public static class ChangeUserType
    {
        [FunctionName("ChangeUserType")]
        public static async Task<HttpResponseMessage> Run([HttpTrigger(AuthorizationLevel.Anonymous, "get", "post", Route = null)] HttpRequestMessage req, TraceWriter log)
        {
            log.Info("C# HTTP trigger function processed a request.");

            // parse query parameter
            string email = req.GetQueryNameValuePairs()
                .FirstOrDefault(q => string.Compare(q.Key, "email", true) == 0)
                .Value;

            if (email == null || email == "")
            {
                // Get request body
                dynamic data = await req.Content.ReadAsAsync<object>();
                email = data?.email;

                if (email == null)
                {
                    return req.CreateResponse(HttpStatusCode.BadRequest, "Error, no email found. ");
                }
            }

            log.Info($"email is {email}");

            var authResult = GetOneAccessToken();
            var graphClient = GetGraphClient(authResult);
            var userId = GetUserID(graphClient, email, log).GetAwaiter().GetResult();

            if (userId == null)
            {
                return req.CreateResponse(HttpStatusCode.BadRequest, $"no user id found for {email}");
            }
            else
            {
                var changeusertype = ChangeGuestUserType(graphClient, log, userId);
                var welcomeGroup = AddUserWelcomeGroup(graphClient, log, userId);

                if (changeusertype != "Success")
                {
                     return req.CreateResponse(HttpStatusCode.BadRequest, $"Usertype was not updated {changeusertype}");
                }

                if (welcomeGroup != "Success")
                {
                    return req.CreateResponse(HttpStatusCode.BadRequest, $"User was not add to the welcome group {welcomeGroup}");
                }
                return req.CreateResponse(HttpStatusCode.OK, $"User {email} has been updated with success");
            }
        }

        public static string GetOneAccessToken()
        {
            string token = "";
            string CLIENT_ID = ConfigurationManager.AppSettings["CLIENT_ID"];
            string CLIENT_SECERET = ConfigurationManager.AppSettings["CLIENT_SECRET"];
            string TENAT_ID = ConfigurationManager.AppSettings["TENANT_ID"];
            string TOKEN_ENDPOINT = "";
            string MS_GRAPH_SCOPE = "";
            string GRANT_TYPE = "";

            try
            {

                TOKEN_ENDPOINT = "https://login.microsoftonline.com/" + TENAT_ID + "/oauth2/v2.0/token";
                MS_GRAPH_SCOPE = "https://graph.microsoft.com/.default";
                GRANT_TYPE = "client_credentials";

            }
            catch (Exception e)
            {
                Console.WriteLine("A error happened while search config file");
            }
            try
            {
                HttpWebRequest request = WebRequest.Create(TOKEN_ENDPOINT) as HttpWebRequest;
                request.Method = "POST";
                request.ContentType = "application/x-www-form-urlencoded";
                StringBuilder data = new StringBuilder();
                data.Append("client_id=" + HttpUtility.UrlEncode(CLIENT_ID));
                data.Append("&scope=" + HttpUtility.UrlEncode(MS_GRAPH_SCOPE));
                data.Append("&client_secret=" + HttpUtility.UrlEncode(CLIENT_SECERET));
                data.Append("&GRANT_TYPE=" + HttpUtility.UrlEncode(GRANT_TYPE));

                byte[] byteData = UTF8Encoding.UTF8.GetBytes(data.ToString());
                request.ContentLength = byteData.Length;
                using (Stream postStream = request.GetRequestStream())
                {
                    postStream.Write(byteData, 0, byteData.Length);
                }

                // Get response

                using (HttpWebResponse response = request.GetResponse() as HttpWebResponse)
                {
                    using (var reader = new StreamReader(response.GetResponseStream()))
                    {
                        JavaScriptSerializer js = new JavaScriptSerializer();
                        var objText = reader.ReadToEnd();
                        LgObject myojb = (LgObject)js.Deserialize(objText, typeof(LgObject));
                        token = myojb.access_token;
                    }
                }
                return token;
            }
            catch (Exception e)
            {
                Console.WriteLine("A error happened while connect to server please check config file");
                return "error";
            }
        }

        public static GraphServiceClient GetGraphClient(string authResult)
        {
            GraphServiceClient graphClient = new GraphServiceClient(
                 new DelegateAuthenticationProvider(
            async (requestMessage) =>
            {
                requestMessage.Headers.Authorization =
                    new AuthenticationHeaderValue("bearer",
                    authResult);
            }));
            return graphClient;
        }

        public static string ChangeGuestUserType(GraphServiceClient graphClient, TraceWriter Log, List<string> userId)
        {
            var message = "";
            var guestUser = new User
            {
                UserType = "Member"
            };
            try
            {
                 graphClient.Users[userId[0]]  //"6c3520af-ddd8-4f77-b18d-44e70b88f4d9"
                .Request()
                .UpdateAsync(guestUser);
                Log.Info($"Change {userId[0]} user type to member successfully");
                message = "Success";
            }
            catch (Exception ex)
            {
                Log.Info($"error message: {ex.Message}");
                message = $"Error: {ex.Message}";
            }
            return message;
        }

        public static async Task<List<string>> GetUserID(GraphServiceClient graphClient, string email, TraceWriter Log)
        {
            List<string> userId = new List<string>();
            try
            {
                var request = await graphClient
                    .Users
                    .Request()
                    .Filter($"userType eq 'guest' and mail eq '{email}'") // apply filter
                    .GetAsync();

                if (request == null)
                {   
                    return null;
                }
                else
                {
                    foreach (var id in request)
                    {
                        userId.Add(id.Id);
                    }
                    return userId;
                }
            }
            catch (ServiceException ex)
            {
                Console.WriteLine($"Error: {ex.Message}");
                return null;
            }
        }
        public static string AddUserWelcomeGroup(GraphServiceClient graphClient, TraceWriter Log, List<string> userId)
        {
            var message = "";
            string welcomeGroup = ConfigurationManager.AppSettings["welcomeGroup"];

            try
            {
                var directoryObject = new DirectoryObject
                {
                    Id = userId[0]
                };

                graphClient.Groups[welcomeGroup].Members.References
                    .Request()
                    .AddAsync(directoryObject);
                Log.Info($" User add to welcomeGroup successfully.");
                message = "Success";
            }
            catch (Exception ex)
            {
                Log.Info($"error message: {ex.Message}");
                message = $"Error: {ex.Message}";
            }
            return message;
        }

    }
}
