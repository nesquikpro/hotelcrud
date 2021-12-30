using Newtonsoft.Json;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Hotel1
{
    public class TypeOfService : ModelAbstract
    {
        public TypeOfService()
        {

        }

        [JsonProperty("IdTypeOfServices")]
        public override int Id { get; set; } = 0;

        public string Name { get; set; } 
        public int Price { get; set; }


        [JsonIgnore]
        public override string Path { get; set; } = "TypeOfServices";
    }
}
