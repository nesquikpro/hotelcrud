using Newtonsoft.Json;
using System;
using System.Collections.Generic;
using System.Net.Http;

namespace Utils
{
    public class AllModel<TObj> where TObj : ModelAbstract
    {
        private string Path { get; set; }

        public AllModel(string path)
        {
            Path = path;
        }

        public List<TObj> Objs { get => Get(); }

        private List<TObj> Get()
        {
            using var client = new HttpClient();
            using var request = new HttpRequestMessage
            {
                Method = HttpMethod.Get,
                RequestUri = new Uri($"http://unknownn-001-site1.gtempurl.com/api/{Path}/"),
            };
            var response = client.Send(request);
            return JsonConvert.DeserializeObject<List<TObj>>(response.Content.ReadAsStringAsync().Result) ??
                new List<TObj>();
        }
    }

}