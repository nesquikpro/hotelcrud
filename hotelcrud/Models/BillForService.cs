﻿using Newtonsoft.Json;
using System;
using Utils;

namespace hotelcrud
{
    public class BillForService : ModelAbstract
    {
        public BillForService()
        {
        }

        [JsonProperty("IdBillForServices")]
        public override int Id { get; set; } = 0;
        public DateTime InvoiceDate { get; set; }
        public int? ClientId { get; set; }
        public int? TypeOfServicesId { get; set; }

        [JsonIgnore]
        public override string Path { get; set; } = "BillForServices";
    }
}
