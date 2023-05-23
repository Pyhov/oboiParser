using System;
using System.Collections.Generic;
using System.Linq;
using System.Net.Http;
using System.Runtime.Remoting.Contexts;
using System.Text;
using System.Threading.Tasks;

namespace oboiParser
{
    public class Product
    {
        public string Path { get; set; } = string.Empty;
        public string Link { get; set; } = string.Empty;
        public string Name { get; set; } = string.Empty;
        public string Price { get; set; } = string.Empty;
        public string Сurrency { get; set; } = string.Empty;
        public string Appointment { get; set; } = string.Empty;
        public string Peculiarities { get; set; } = string.Empty;
        public string ContentsOfDelivery { get; set; } = string.Empty;
        public string OuterHtmlCharacteristics { get; set; } = string.Empty;
        public string Documentation { get; set; } = string.Empty;
        public bool IsTableCharacteristics { get; set; }  = false;
        public Dictionary<string,string> Characteristics { get; set; } = new Dictionary<string, string>();
        public List<string> ImageUrls { get; set; } = new List<string>();
    }

    public class Characteristic
    {
        public Characteristic(string name, string value)
        {
            Name = name;
            Value = value;
        }

        public string Name { get; set; } = string.Empty;
        public string Value { get; set; } = string.Empty;
       
    }
}
