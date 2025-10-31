using Newtonsoft.Json;
using Utils;

namespace hotelcrud
{
    public class Post : ModelAbstract
    {
        public Post()
        {

        }

        [JsonProperty("IdPost")]
        public override int Id { get; set; } = 0;

        public string Name { get; set; } 
        public int Salary { get; set; }

        [JsonIgnore]
        public override string Path { get; set; } = "Posts";
    }
}
