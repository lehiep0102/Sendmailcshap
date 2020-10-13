using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using RestSharp;
using Newtonsoft.Json;
using System.Windows.Forms;
using System.IO;
namespace MAS_EMAIL
{
    class HelloWorld
    {
 
        public string login()
        {
            string link;
            var client = new RestClient("http://10.0.25.5");
            var request = new RestRequest("/api/auth/login", Method.POST);
            request.AddHeader("Content-Type", "application/json");
            request.AddParameter("username", "admin");
            request.AddParameter("password", "123456789");
            var restResponse = client.Execute(request);
            dynamic document = JsonConvert.DeserializeObject(restResponse.Content);
            link = document.data.token;
            return link;
        }

        public string send(string authen,string file )
        {
            string link;
            var client = new RestClient("http://10.0.25.5");
            //client.Timeout = -1;
            var request = new RestRequest("/api/common/upload",Method.POST);
            request.AddHeader("Authorization", "Bearer "+ authen);
            request.AddFile("file", file);
            request.AddParameter("type", "contract");
            IRestResponse response = client.Execute(request);
            dynamic document = JsonConvert.DeserializeObject(response.Content);
            link = document.data.uri;
            return link;
            //MessageBox.Show(link, "aa");
        }
    }
}
