using Newtonsoft.Json;
using Utils;

namespace hotelcrud
{
    public class RoomNumber : ModelAbstract
    {
        public RoomNumber()
        {

        }
        [JsonProperty("IdNumber")]
        public override int Id { get; set; } = 0;
        public int RoomNumber1 { get; set; }
        public int RoomFloor { get; set; }
        public int Price { get; set; }

        [JsonIgnore]
        public override string Path { get; set; } = "RoomNumbers";
    }
}
