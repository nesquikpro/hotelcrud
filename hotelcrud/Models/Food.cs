﻿using Newtonsoft.Json;
using Utils;

namespace hotelcrud
{
    public class Food : ModelAbstract
    {
        public Food()
        {

        }

        [JsonProperty("IdFood")]
        public override int Id { get; set; } = 0;

        public string Name { get; set; } 
        public int Price { get; set; }

        [JsonIgnore]
        public override string Path { get; set; } = "Foods";
    }
}

