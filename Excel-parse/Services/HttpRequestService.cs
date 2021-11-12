using System;
using System.Collections;
using System.Collections.Generic;
using System.Net.Http;
using System.Net.Http.Headers;
using System.Text;
using System.Threading.Tasks;

namespace Excel_parse.Services
{
    public class HttpRequestService
    {
        public HttpRequestService()
        {

        }

        public HttpResponseMessage PostDataAsJsonAsync(string url, PostData postData)
        {
            try
            {
                using (var httpClient = new HttpClient())
                {
                    //var token = "eyJhbGciOiJodHRwOi8vd3d3LnczLm9yZy8yMDAxLzA0L3htbGRzaWctbW9yZSNyc2Etc2hhMjU2IiwidHlwIjoiSldUIn0.eyJ0ZW5hbnRfaWQiOiJENDQxRkYyNy1ERTUxLTQxREYtOEFDNi02NUEyNDM4RjYyMUEiLCJzdWIiOiI0OTk3ODZjOS05M2JlLTQ2ZTctODhlNi02MGFjMWYyZGUzYzAiLCJzaXRlX2lkIjoiRDQ0M0EyOTQtQTEwQS00MjM0LUFDMjUtRjBCRjFBQjFDRDQ0Iiwib3JpZ2luIjoic3RnLXN3aWNhLXBvcnRhbC5zZWxpc2UuYml6Iiwic2Vzc2lvbl9pZCI6IjgwZTJjMmNhODJkYjRiNzViMzc0YjA4ODJjZTg1MDg5IiwidXNlcl9pZCI6IjQ5OTc4NmM5LTkzYmUtNDZlNy04OGU2LTYwYWMxZjJkZTNjMCIsImRpc3BsYXlfbmFtZSI6IkNsaWVudCBDcmVkZW50YWlscyIsInNpdGVfbmFtZSI6IkVjYXAgVGVhbSIsInVzZXJfbmFtZSI6IjQ5OTc4NmM5LTkzYmUtNDZlNy04OGU2LTYwYWMxZjJkZTNjMCIsImVtYWlsIjoiaW5mb0BzZWxpc2UuY2giLCJwaG9uZV9udW1iZXIiOiIiLCJsYW5ndWFnZSI6IkVOIiwidXNlcl9sb2dnZWRpbiI6IlRydWUiLCJuYW1lIjoiNDk5Nzg2YzktOTNiZS00NmU3LTg4ZTYtNjBhYzFmMmRlM2MwIiwicm9sZSI6WyJhZG1pbiIsImFwcHVzZXIiXSwibmJmIjoxNTU1MzkzNjA1LCJleHAiOjE1NzA1NzM2MDUsImlzcyI6IkNOPUVudGVycHJpc2UgQ2xvdWQgQXBwbGljYXRpb24gUGxhdGZvcm0iLCJhdWQiOiIqIn0.ZXH0uTo4wifGLF_EPjXd2aowXfo35_Nvg5pa3wo9YkUkCGTNHgd4xrRrmjkW0KezoUsvPGikCf8b9wtbv9xbVejTrub1v0yf9T6C1irpKyq0Qwmb7nGUcfz5yFLChFrKOR9o_Xse8mbc16alR7fZz8GyMrbmjx8JHrTNvHxIP__OZgU5xrl5OWetQGSmJ7jxa-NaNKNqbSjttgB2hZoy3wTzEzVa6pf5aoZUupbht3KDoYOhdD4Y-_RAc_vsJMcpbPDEkD40qSTkIGJFWqAxN9IdvqQQK9OgihZDesmLzS3KW4ZZoP2QQ1U2_uQgGHVsTXcF-Nlod4E_QUpZLRILOg";
                    var token = "eyJhbGciOiJSUzI1NiIsInR5cCI6IkpXVCJ9.eyJzdWIiOiIyM2UxYmNlYy1mNTk1LTRmNjEtYTAyNy01NjE4YTVmNmIwMzMiLCJ0ZW5hbnRfaWQiOiJENDQxRkYyNy1ERTUxLTQxREYtOEFDNi02NUEyNDM4RjYyMUEiLCJzaXRlX2lkIjoiRDQ0M0EyOTQtQTEwQS00MjM0LUFDMjUtRjBCRjFBQjFDRDQ0Iiwib3JpZ2luIjoiYmV0YS5zd2ljYS1ndWlkZS5jaCIsInVzZXJfaWQiOiIyM2UxYmNlYy1mNTk1LTRmNjEtYTAyNy01NjE4YTVmNmIwMzMiLCJkaXNwbGF5X25hbWUiOiJTd2ljYSBUaGVyYXBpc3QiLCJzaXRlX25hbWUiOiJTd2ljYSIsInVzZXJfbmFtZSI6IjQyOWIyMzhmLTA3NjctNGI4Yy1hN2YwLTRjMGJlN2ZmMTk5MyIsImVtYWlsIjoic3dpY2EudGhlcmFwaXN0QGdtYWlsLmNvbSIsInVzZXJfbG9nZ2VkaW4iOiJUcnVlIiwic2Vzc2lvbl9pZCI6ImMxOWI4NjliLTZmNzEtNGNiOS05ZWUwLThmMDBmZTkzMDBkNCIsImxhbmd1YWdlIjoiZW4tVVMiLCJwaG9uZV9udW1iZXIiOiIrODgwMTcxODA2MDA2OSIsInJvbGUiOlsiYXBwdXNlciIsImFub255bW91cyIsInRoZXJhcGlzdCJdLCJuYmYiOjE1NTU1NzEzODAsImV4cCI6MTU1NTYwNzM4MCwiaXNzIjoiQ049RW50ZXJwcmlzZSBDbG91ZCBBcHBsaWNhdGlvbiBQbGF0Zm9ybSIsImF1ZCI6IioifQ.cfwY_E3TscXFdZRYw8UyYT0gF5E8VXXl3FSbpt0j_aam4v45xpz1cCb49j7LVGUEpBFjAdMuZqRZl4sTWh2O41BJ-E6sV7S4R4-wPqAx4Jm6wNaW1TRJQ6j1MZpPYFB7OfwjXtZcIWT8kY42KOYJdwzy0JpNOgTo_TOzVBn3Vf_JPVPsTk7ZOQ5HWIW3lX_0JVwut85lapFYA_34vUeY9WnzDsjovHXeU8yh6nbXNBmwuIZ4psSgzQMA8gnlRStfLQQz3mv0feRBsKTVn-2xazyHFjW-JgGQON4W2V7oHBdfGS2pGgOPDBIrbH2TTT8NX2zSKgq3nki2upvuG-5oIQ";
                    httpClient.DefaultRequestHeaders.Add("Origin", "https://beta.swica-guide.ch/");
                    httpClient.DefaultRequestHeaders.Authorization = new AuthenticationHeaderValue("bearer", token);

                    ArrayList arrayList = new ArrayList();
                    arrayList.Add(postData);
                    var jsonData = Newtonsoft.Json.JsonConvert.SerializeObject(arrayList);

                    var content = new StringContent(jsonData, Encoding.UTF8, "application/json");
                    //if (!string.IsNullOrEmpty(jsonData))
                    //{
                    //    content = new StringContent(jsonData, Encoding.UTF8, "text/plain");
                    //}

                    return httpClient.PostAsync(url, content).Result;
                    //var response = httpClient.PostAsync(url, content);

                    //if (response.IsSuccessStatusCode)
                    //{
                    //    return response;
                    //}

                    //return response;
                }
            }
            catch (Exception ex)
            {
                //throw;
            }

            return null;
        }

    }
}
