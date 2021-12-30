using Newtonsoft.Json;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Hotel1
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
