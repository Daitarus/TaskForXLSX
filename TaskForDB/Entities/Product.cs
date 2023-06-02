using System;
using System.Collections.Generic;
using System.Linq;
using System.Net;
using System.Text;
using System.Threading.Tasks;

namespace TaskForDB.Entities
{
    public class Product : Entity
    {
        public string Name { get; set; }
        public string NameUnit { get; set; }
        public double Price { get; set; }

        public Product(uint id, string sheetName) : base(id, sheetName) { }
        public Product(uint id, string sheetName, string name, string nameUnit, string price) : base(id, sheetName)
        {
            try
            {
                Price = double.Parse(price);
            }
            catch(Exception ex)
            {
                throw new ArgumentException("Argument cannot be converted to double", nameof(price), ex);
            }
            Name = name;
            NameUnit = nameUnit;
        }
        public Product(uint id, string sheetName, string name, string nameUnit, double price) : base(id, sheetName)
        {
            Name = name;
            NameUnit = nameUnit;
            Price = price;
        }
    }
}
