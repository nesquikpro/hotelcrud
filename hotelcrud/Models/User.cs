using Newtonsoft.Json;
using Utils;

namespace hotelcrud
{
    public class User : ModelAbstract
    {
        public User()
        {

        }

        [JsonProperty("IdUser")]
        public override int Id { get; set; } = 0;

        public string Login { get; set; } 
        public string Password { get; set; }
        public int? RoleId { get; set; }

        [JsonIgnore]
        public override string Path { get; set; } = "Users";
    }
}
