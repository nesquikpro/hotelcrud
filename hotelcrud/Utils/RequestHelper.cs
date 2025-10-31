using Newtonsoft.Json;
using System;
using System.Net.Http;
using System.Text;
using System.Threading.Tasks;

namespace Utils
{
    public static class RequestHelper
    {
        private const string ApiPath = "http://unknownn-001-site1.gtempurl.com/api/";

        public static async Task<TObj> Get<TObj>(this TObj obj) where TObj : ModelAbstract
        {
            using var client = new HttpClient();
            using var request = new HttpRequestMessage
            {
                Method = HttpMethod.Get,
                RequestUri = new Uri($"{ApiPath}{obj.Path}/{obj.Id}")
            };
            var response = await client.SendAsync(request);
            return JsonConvert.DeserializeObject<TObj>(response.Content.ReadAsStringAsync().Result) ??
            default!;
        }

        public static async Task<TObj> Update<TObj>(this TObj obj) where TObj : ModelAbstract
        {
            using var client = new HttpClient();
            using var request = new HttpRequestMessage
            {
                Method = HttpMethod.Put,
                RequestUri = new Uri($"{ApiPath}{obj.Path}/{obj.Id}"),
                Content =
            new StringContent(JsonConvert.SerializeObject(obj), Encoding.UTF8, "application/json"),
            };
            var response = await client.SendAsync(request);
            return JsonConvert.DeserializeObject<TObj>(response.Content.ReadAsStringAsync().Result) ??
            default!;
        }

        public static async Task<TObj> Add<TObj>(this TObj obj) where TObj : ModelAbstract
        {
            using var client = new HttpClient();
            using var request = new HttpRequestMessage
            {
                Method = HttpMethod.Post,
                RequestUri = new Uri($"{ApiPath}{obj.Path}/"),
                Content =
            new StringContent(obj.ToJson(), Encoding.UTF8, "application/json"),
            };
            var response = await client.SendAsync(request);
            var fdf = response.Content.ReadAsStringAsync().Result;
            return JsonConvert.DeserializeObject<TObj>(response.Content.ReadAsStringAsync().Result) ??
            default!;
        }

        public static async Task<bool> Delete<TObj>(this TObj obj) where TObj : ModelAbstract
        {
            using var client = new HttpClient();
            using var request = new HttpRequestMessage
            {
                Method = HttpMethod.Delete,
                RequestUri = new Uri($"{ApiPath}{obj.Path}/{obj.Id}")
            };
            var response = await client.SendAsync(request);
            return response.IsSuccessStatusCode;
        }
    }
}