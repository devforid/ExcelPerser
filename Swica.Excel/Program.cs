using Excel_parse.Services;
using System;
using System.Collections;
using System.Collections.Generic;
using System.Dynamic;
using System.Net.Http;
using System.Net.Http.Headers;
using System.Text;
using System.Threading.Tasks;

namespace Swica.Excel
{
    class Program
    {
        private static readonly HttpClient client = new HttpClient();
        public static void Main(string[] args)
        {
            Console.WriteLine("Hello World!");
            //HttpResponseMessage response = PostAsJsonAsync(requesturl.ToString(), postData);


            ExcelParserService excelParserService = new ExcelParserService();
            excelParserService.ParseExcelFile();
        }
 

        //private static HttpResponseMessage PostAsJsonAsync(string url, dynamic postData)
        //{
        //    try
        //    {
        //        using (var httpClient = new HttpClient())
        //        {
        //            var token = "eyJhbGciOiJSUzI1NiIsInR5cCI6IkpXVCJ9.eyJ0ZW5hbnRfaWQiOiJENDQxRkYyNy1ERTUxLTQxREYtOEFDNi02NUEyNDM4RjYyMUEiLCJzdWIiOiI3ZDhiMjFlNi1mNjBmLTQwZTAtOGUxZi0xODE3ZTRlM2YyNTEiLCJzaXRlX2lkIjoiRDQ0M0EyOTQtQTEwQS00MjM0LUFDMjUtRjBCRjFBQjFDRDQ0Iiwib3JpZ2luIjoic3RnLXN3aWNhLXBvcnRhbC5zZWxpc2UuYml6Iiwic2Vzc2lvbl9pZCI6ImVjYXAtYzgxNjNiNzgtMGFkYS00MjViLWE2OGYtMzZjNzVkZGMwYWNkIiwidXNlcl9pZCI6IjdkOGIyMWU2LWY2MGYtNDBlMC04ZTFmLTE4MTdlNGUzZjI1MSIsImRpc3BsYXlfbmFtZSI6Ikp1YmF5ZXIgQWwgRmFyYWJpIiwic2l0ZV9uYW1lIjoiRWNhcCBUZWFtIiwidXNlcl9uYW1lIjoiN2Q4YjIxZTYtZjYwZi00MGUwLThlMWYtMTgxN2U0ZTNmMjUxIiwiZW1haWwiOiJqdWJheWVyLmZhcmFiaUBzZWxpc2UuY2giLCJwaG9uZV9udW1iZXIiOiIrODgwMTc0NDIyMTQyNSIsImxhbmd1YWdlIjoiZW4tVVMiLCJ1c2VyX2xvZ2dlZGluIjoiVHJ1ZSIsIm5hbWUiOiI3ZDhiMjFlNi1mNjBmLTQwZTAtOGUxZi0xODE3ZTRlM2YyNTEiLCJ1c2VyX2F1dG9fZXhwaXJlIjoiRmFsc2UiLCJ1c2VyX2V4cGlyZV9vbiI6IjAxLzAxLzAwMDEgMDA6MDA6MDAiLCJyb2xlIjpbInRoZXJhcGlzdCIsImFwcHVzZXIiLCJhbm9ueW1vdXMiXSwibmJmIjoxNTU1MzMwNjM5LCJleHAiOjE1NTUzMzEwNTksImlzcyI6IkNOPUVudGVycHJpc2UgQ2xvdWQgQXBwbGljYXRpb24gUGxhdGZvcm0iLCJhdWQiOiIqIn0.XBaSudQTBLXI07tAHW-7ST4JANCQBU5k9tD7KnZ8Vy7zMJLSNJRv5jho40PVNkANDj6o9gr4yV20D_y5Us01cISGHlAeqlWNXgJsVM6pFG55YIlUDtaI_s59zyxHUCpvqv7fzKD1TB5UO-MlyFzTSrwFfRhHCGDoHd1bZFjawS1p_G1YNQfyjCd3vI_dsYCzVtIfTA-PwdhiBI5H2uZXga4T0gQPjt4trzSjY1f3PPKiQDCaf3nwCCb_jpdRCG2tcJs1kxefjqZAYpoVCpmmpi0Fm7k34V2pI__Te0rWcQK5HqeAKtHugws9pMnQvYq0CZyCZ0Zj8gZC64q5bGEpsQ";
        //            httpClient.DefaultRequestHeaders.Add("Origin", "https://stg-swica-portal.selise.biz/");
        //            httpClient.DefaultRequestHeaders.Authorization = new AuthenticationHeaderValue("bearer", token);

        //            //foreach (var header in _requestInfo.Headers)
        //            //{
        //            //    if (!header.Key.Equals("Authorization"))
        //            //    {
        //            //        httpClient.DefaultRequestHeaders.Add(header.Key, header.Value);
        //            //    }
        //            //    if (header.Key.Equals("Host"))
        //            //    {
        //            //        var serviceUri = new Uri(url);
        //            //        httpClient.DefaultRequestHeaders.Host = serviceUri.Host;
        //            //    }
        //            //}

        //            //if (additionalHeaders != null)
        //            //{
        //            //    foreach (var header in additionalHeaders)
        //            //    {
        //            //        httpClient.DefaultRequestHeaders.Add(header.Key, header.Value);
        //            //    }
        //            //}
        //            ArrayList arrayList = new ArrayList();
        //            arrayList.Add(postData);
        //            var jsonData = Newtonsoft.Json.JsonConvert.SerializeObject(arrayList);

        //            var content = new StringContent(jsonData, Encoding.UTF8, "application/json");
        //            //if (!string.IsNullOrEmpty(jsonData))
        //            //{
        //            //    content = new StringContent(jsonData, Encoding.UTF8, "text/plain");
        //            //}

        //            var response = httpClient.PostAsync(url, content).Result;

        //            if (response.IsSuccessStatusCode)
        //            {
        //                return response;
        //            }
        //            //else if (useImpersonation && response.StatusCode == HttpStatusCode.Unauthorized)
        //            //{
        //            //    _logger.Info("Found Unauthorized request. Sending request through impersonation");
        //            //    httpClient.DefaultRequestHeaders.Authorization = new AuthenticationHeaderValue(_requestInfo.AuthorizationScheme, GetImpersonatedAccessToken());
        //            //    response = httpClient.PostAsync(url, content).Result;
        //            //    if (response.IsSuccessStatusCode)
        //            //    {
        //            //        return response;
        //            //    }
        //            //}

        //            //_logger.Info("Error Uri :: " + response.RequestMessage.RequestUri.ToString());
        //            //_logger.Info("Error StatusCode :: " + response.StatusCode.ToString());
        //            //_logger.Info("Error response :: " + response.ToString());

        //            return response;
        //        }
        //    }
        //    catch (Exception ex)
        //    {
        //        throw;
        //    }
        //}


    }
}
