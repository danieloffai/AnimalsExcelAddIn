using Newtonsoft.Json;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Net.Http;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace ExcelAddInZoo
{
    public static class APIsession
    {
        public static List<ZooAnimal> GetAnimals(int animalNumber)
        {
            List<ZooAnimal> animals = new List<ZooAnimal>();

            using (HttpClient client = new HttpClient())
            {
                try
                {
                    string url = $@"http://zoo-animal-api.herokuapp.com/animals/rand/{animalNumber}";
                    HttpResponseMessage response = client.GetAsync(url).Result;
                    animals = JsonConvert.DeserializeObject<List<ZooAnimal>>(response.Content.ReadAsStringAsync().Result);
                }
                catch (Exception ex)
                {
                    MessageBox.Show(ex.Message);
                }
            }

            return animals;
        }
    }
}
