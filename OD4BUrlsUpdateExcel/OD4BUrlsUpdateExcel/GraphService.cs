
using System;
using System.Collections.Generic;
using System.Data;
using System.IO;
using System.Linq;
using System.Net;
using System.Net.Http;
using System.Net.Http.Headers;
using System.Text;
using System.Threading.Tasks;
using System.Globalization;
using System.Configuration;
using Newtonsoft.Json.Linq;
using Newtonsoft.Json;

namespace OD4BUrlsUpdateExcel
{
    public class GraphService
    {

        static private string redirectUri = "http://localhost:55065/";
        static private string appId =  "";
       
        static private string appSecret = "";

        

        //private ConnectionStringSettings connectionStringSettings;

        static private string scopes = "Files.Read.All Groups.Read.All";
        public async Task<string> GetOD4BUrl(string userAlias)
        {
            string OD4BUrl = string.Empty;
            string accessToken = await GetTokenForAPP();
            string endpoint = "https://graph.microsoft.com/v1.0/users/" + userAlias.Trim() + "@microsoft.com/drive/root";
            string queryParameter = "?$select=WebUrl";
            var user = await CallMSGraphAPI(endpoint, queryParameter, accessToken);
            if (user.GetValue("webUrl") != null)
            {
                OD4BUrl = !string.IsNullOrEmpty(user.GetValue("webUrl").ToString()) ? user.GetValue("webUrl").ToString() : string.Empty;
                OD4BUrl = OD4BUrl.Substring(0, OD4BUrl.Length - 10);
            }
            return OD4BUrl;


        }

        public async Task<string> GetGroupsUrl(string userAlias)
        {
            DataTable dtGroups = new DataTable();
            string groupUrls = string.Empty;
            //List<Group> groups = new List<Group>();
            string accessToken = await GetTokenForAPP();
            string endpoint = "https://graph.microsoft.com/beta/users/" + userAlias.Trim() + "@microsoft.com/joinedgroups";
            string queryParameter = string.Empty;
            var json = await CallMSGraphAPI(endpoint, queryParameter, accessToken);
            if (json.GetValue("value") != null)
            {
                string strJ = json.GetValue("value").ToString();
                dtGroups = (DataTable)JsonConvert.DeserializeObject(strJ, (typeof(DataTable)));
                if (dtGroups.AsEnumerable().Count() > 0)
                {
                    foreach (var row in dtGroups.AsEnumerable())
                    {

                        string driveEndpoint = "https://graph.microsoft.com/v1.0/groups/" + row["id"].ToString() + "/drive/root";
                        string driveQuery = "?$select=WebUrl";
                        var driveDetails = await CallMSGraphAPI(driveEndpoint, driveQuery, accessToken);
                        if (driveDetails.GetValue("webUrl") != null)
                        {
                            string groupDriveUrl = !string.IsNullOrEmpty(driveDetails.GetValue("webUrl").ToString()) ? driveDetails.GetValue("webUrl").ToString() : string.Empty;
                            groupUrls = groupUrls + ";" + groupDriveUrl;
                        }
                    }
                }

            }

            return groupUrls;

        }

        public async Task<JObject> CallMSGraphAPI(string endpoint, string queryParameter, string accessToken)
        {
            JObject json = new JObject();
            using (var client = new HttpClient())
            {
                using (var request = new HttpRequestMessage(HttpMethod.Get, endpoint + queryParameter))
                {
                    request.Headers.Accept.Add(new MediaTypeWithQualityHeaderValue("application/json"));
                    request.Headers.Authorization = new AuthenticationHeaderValue("Bearer", accessToken);

                    // This header has been added to identify our sample in the Microsoft Graph service. If extracting this code for your project please remove.
                    request.Headers.Add("SampleID", "aspnet-connect-rest-sample");

                    using (var response = await client.SendAsync(request))
                    {
                        if (response.IsSuccessStatusCode)
                        {
                            json = JObject.Parse(await response.Content.ReadAsStringAsync());
                            //string strJ = json.GetValue("value").ToString();
                            //dtGroups = (DataTable)JsonConvert.DeserializeObject(strJ, (typeof(DataTable)));
                            //string add = !string.IsNullOrEmpty(json.GetValue("mail").ToString()) ? json.GetValue("mail").ToString() : json.GetValue("userPrincipalName").ToString();
                            //TraceManager.CreateInstance("MDLWeb").LogInformation(EventLogID.MDLSystemMaintenanceInformation,
                            //    string.Format(CultureInfo.InvariantCulture, "{0} for the MS Graph endpoint: {1} ", response.ReasonPhrase, endpoint.ToString(CultureInfo.InvariantCulture))
                            //    , "MDLWeb");
                           
                        }
                        else
                        {
                            //TraceManager.CreateInstance("MDLWeb").LogWarning(EventLogID.MDLSystemMaintenanceInformation,
                            //    string.Format(CultureInfo.InvariantCulture, "{0} for the MS Graph endpoint: {1} ", response.ReasonPhrase, endpoint.ToString(CultureInfo.InvariantCulture))
                            //    , "MDLWeb");
                           
                        }
                        //return me.Address?.Trim();
                    }
                }
            }
            return json;
        }

        public async Task<string> GetTokenForAPP()
        {

            //  string html = string.Empty;
            //  string url = @"https://login.microsoftonline.com/microsoft.onmicrosoft.com/adminconsent?client_id=" + appId + "&state=12345&redirect_uri=" + redirectUri;

            ////  https://login.microsoftonline.com/microsoft.onmicrosoft.com/adminconsent?client_id=904585ee-0c38-4f4a-a159-f45b3cde546b&state=12345&redirect_uri=http://localhost:55065

            //  HttpWebRequest request = (HttpWebRequest)WebRequest.Create(url);
            //  //request.AutomaticDecompression = DecompressionMethods.GZip;

            //  using (HttpWebResponse response = (HttpWebResponse)request.GetResponse())
            //  using (Stream stream = response.GetResponseStream())
            //  using (StreamReader reader = new StreamReader(stream))
            //  {
            //      html = reader.ReadToEnd();
            //  }

            //  Console.WriteLine(html);
            //appId = Paf.Config.GetAppSetting("MSGAppID");
            string postData = "client_id=" + appId + "&scope=https://graph.microsoft.com/.default&client_secret=" + appSecret + "&grant_type=client_credentials";
            HttpWebRequest myHttpWebRequest = (HttpWebRequest)HttpWebRequest.Create("https://login.microsoftonline.com/microsoft.onmicrosoft.com/oauth2/v2.0/token");

            myHttpWebRequest.Method = "POST";

            byte[] data = Encoding.ASCII.GetBytes(postData);

            myHttpWebRequest.ContentType = "application/x-www-form-urlencoded";
            myHttpWebRequest.ContentLength = data.Length;

            Stream requestStream = myHttpWebRequest.GetRequestStream();
            requestStream.Write(data, 0, data.Length);
            requestStream.Close();

            HttpWebResponse myHttpWebResponse = (HttpWebResponse)myHttpWebRequest.GetResponse();

            Stream responseStream = myHttpWebResponse.GetResponseStream();

            StreamReader myStreamReader = new StreamReader(responseStream, Encoding.Default);

            string pageContent = myStreamReader.ReadToEnd();

            myStreamReader.Close();
            responseStream.Close();

            myHttpWebResponse.Close();


            var json = JObject.Parse(pageContent);

            return json.GetValue("access_token").ToString();
        }
    }
}
