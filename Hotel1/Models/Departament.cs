using Newtonsoft.Json;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Hotel1
{
    public class Departament :ModelAbstract
    {
        public Departament()
        {

        }

        [JsonProperty("IdDepartament")]
        public override int Id { get; set; } = 0;

        public string Name { get; set; } 
        public string DepartamentPhone { get; set; }

        [JsonIgnore]
        public override string Path { get; set; } = "Departaments";
    }
}
