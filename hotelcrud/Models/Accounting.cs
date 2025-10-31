using Newtonsoft.Json;
using System;
using Utils;

namespace hotelcrud
{
    public class Accounting : ModelAbstract
    {
        public Accounting()
        {

        }

        [JsonProperty("IdAccounting")]
        public override int Id { get; set; } = 0;

        public int Amount { get; set; }
        public DateTime DateIssue { get; set; }
        public int? EmployeeId { get; set; }

        [JsonIgnore]
        public override string Path { get; set; } = "Accountings";
    }
}
