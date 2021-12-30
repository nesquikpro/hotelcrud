using Newtonsoft.Json;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Hotel1
{
    public class OrderingFood : ModelAbstract
    {
        public OrderingFood()
        {

        }

        [JsonProperty("IdOrderingFood")]
        public override int Id { get; set; } = 0;

        public int? FoodId { get; set; }
        public int? ClientId { get; set; }

        [JsonIgnore]
        public override string Path { get; set; } = "OrderingFoods";
    }
}
