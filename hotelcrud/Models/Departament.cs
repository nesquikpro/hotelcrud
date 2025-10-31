using Newtonsoft.Json;
using Utils;

namespace hotelcrud
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
