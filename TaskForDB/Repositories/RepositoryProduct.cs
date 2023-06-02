using ClosedXML.Excel;
using DocumentFormat.OpenXml.Office2010.Excel;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using TaskForDB.Entities;

namespace TaskForDB.Repositories
{
    internal class RepositoryProduct : Repository<Product>
    {
        public RepositoryProduct(IXLWorksheet worksheet) : base(worksheet) { }

        public override List<Product> SelectAll()
        {
            List<Product> listOfProducts = new List<Product>();

            var rows = worksheet.RangeUsed().RowsUsed().ToList();
            for (int i = 1; i < rows.Count(); i++)
            {
                var row = rows[i];
                try
                {
                    string sheetName = worksheet.Name;
                    uint id = uint.Parse(row.Cell(1).Value.ToString());
                    string name = row.Cell(2).Value.ToString();
                    string nameUnit = row.Cell(3).Value.ToString();
                    double price = double.Parse(row.Cell(4).Value.ToString());
                    Product product = new Product(id, sheetName, name, nameUnit, price);
                    listOfProducts.Add(product);

                }
                catch (Exception ex)
                {
                    throw new Exception($"Unable to convert incoming strings to {nameof(Product)}", ex);
                }
            }

            return listOfProducts;
        }
        public override Product? GetForId(uint id)
        {
            var rows = worksheet.RangeUsed().RowsUsed().ToList();
            for (int i = 1; i < rows.Count(); i++)
            {
                var row = rows[i];
                try
                {
                    if (id == uint.Parse(row.Cell(1).Value.ToString()))
                    {
                        string sheetName = worksheet.Name;
                        string name = row.Cell(2).Value.ToString();
                        string nameUnit = row.Cell(3).Value.ToString();
                        double price = double.Parse(row.Cell(4).Value.ToString());
                        return new Product(id, sheetName, name, nameUnit, price);
                    }

                }
                catch (Exception ex)
                {
                    throw new Exception($"Unable to convert incoming strings to {nameof(Product)}", ex);
                }
            }
            return null;
        }
        public Product? GetForName(string name)
        {
            var rows = worksheet.RangeUsed().RowsUsed().ToList();
            for (int i = 1; i < rows.Count(); i++)
            {
                var row = rows[i];
                try
                {
                    if (name == row.Cell(2).Value.ToString())
                    {
                        string sheetName = worksheet.Name;
                        uint id = uint.Parse(row.Cell(1).Value.ToString());
                        string nameUnit = row.Cell(3).Value.ToString();
                        double price = double.Parse(row.Cell(4).Value.ToString());
                        return new Product(id, sheetName, name, nameUnit, price);
                    }

                }
                catch (Exception ex)
                {
                    throw new Exception($"Unable to convert incoming strings to {nameof(Product)}", ex);
                }
            }
            return null;
        }
    }
}
