using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ExcelAddInZoo
{
    public class ZooAnimal
    {
        public string name { get; set; }
        public string latin_name { get; set; }
        public string animal_type { get; set; }
        public string active_time { get; set; }
        public double length_min { get; set; }
        public double length_max { get; set; }
        public double weight_min { get; set; }
        public double weight_max { get; set; }
        public double lifespan { get; set; }
        public string habitat { get; set; }
        public string diet { get; set; }
        public string geo_range { get; set; }
        public int id { get; set; }
    }
}
