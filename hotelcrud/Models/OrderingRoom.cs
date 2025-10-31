using Newtonsoft.Json;
using System;
using Utils;

namespace hotelcrud
{
    public class OrderingRoom : ModelAbstract
    {
        public OrderingRoom()
        {

        }

        [JsonProperty("IdOrderingRoom")]
        public override int Id { get; set; } = 0;
        public DateTime ArrivalDate { get; set; }
        public DateTime DepartureDate { get; set; }
        public int? ClientId { get; set; }
        public int? EmployeeId { get; set; }
        public int? NumberId { get; set; }

        [JsonIgnore]
        public override string Path { get; set; } = "OrderingRooms";
    }
}
