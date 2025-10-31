using Newtonsoft.Json;
using Utils;

namespace hotelcrud
{
    public class Role : ModelAbstract
    {
        public Role()
        {

        }

        [JsonProperty("IdRole")]
        public override int Id { get; set; } = 0;
        public string Name { get; set; }

        [JsonIgnore]
        public override string Path { get; set; } = "Roles";
    }
}
