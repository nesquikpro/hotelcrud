using Newtonsoft.Json;
using System;

namespace Hotel1
{
    public class Employee : ModelAbstract
    {
        public Employee()
        {

        }

        [JsonProperty("IdEmployee")]
        public override int Id { get; set; } = 0;

        public string Name { get; set; } 
        public string Surname { get; set; } 
        public string MiddleName { get; set; } 
        public string SeriaPasport { get; set; } 
        public string NumberPasport { get; set; } 
        public string EmployeePhone { get; set; } 
        public DateTime DateBirthEmployee { get; set; }
        public string EmployeeEmail { get; set; } 
        public int? PostId { get; set; }
        public int? DepartamentId { get; set; }
        public int? UserId { get; set; }


        [JsonIgnore]
        public override string Path { get; set; } = "Employees";
    }
}

