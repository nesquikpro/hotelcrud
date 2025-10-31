using Newtonsoft.Json;
using Utils;

namespace hotelcrud
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
