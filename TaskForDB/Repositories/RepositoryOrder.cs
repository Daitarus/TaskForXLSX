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
    internal class RepositoryOrder : Repository<Order>
    {
        public RepositoryOrder(IXLWorksheet worksheet) : base(worksheet) { }

        public override List<Order> SelectAll()
        {
            List<Order> listOfOrders = new List<Order>();

            var rows = worksheet.RangeUsed().RowsUsed().ToList();
            for (int i = 1; i < rows.Count(); i++)
            {
                var row = rows[i];
                try
                {
                    string sheetName = worksheet.Name;
                    uint id = uint.Parse(row.Cell(1).Value.ToString());
                    uint productId = uint.Parse(row.Cell(2).Value.ToString());
                    uint clientId = uint.Parse(row.Cell(3).Value.ToString());
                    uint numberOfOrder = uint.Parse(row.Cell(4).Value.ToString());
                    uint amount = uint.Parse(row.Cell(5).Value.ToString());
                    DateTime dateOfCreation = DateTime.Parse(row.Cell(6).Value.ToString());
                    Order order = new Order(id, sheetName, productId, clientId, numberOfOrder, amount, dateOfCreation);
                    listOfOrders.Add(order);

                }
                catch (Exception ex)
                {
                    throw new Exception($"Unable to convert incoming strings to {nameof(Order)}", ex);
                }
            }

            return listOfOrders;
        }
        public override Order? GetForId(uint id)
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
                        uint productId = uint.Parse(row.Cell(2).Value.ToString());
                        uint clientId = uint.Parse(row.Cell(3).Value.ToString());
                        uint numberOfOrder = uint.Parse(row.Cell(4).Value.ToString());
                        uint amount = uint.Parse(row.Cell(5).Value.ToString());
                        DateTime dateOfCreation = DateTime.Parse(row.Cell(6).Value.ToString());
                        return new Order(id, sheetName, productId, clientId, numberOfOrder, amount, dateOfCreation);
                    }

                }
                catch (Exception ex)
                {
                    throw new Exception($"Unable to convert incoming strings to {nameof(Order)}", ex);
                }
            }
            return null;
        }

        public List<Order> SelectForProductId(uint productId)
        {
            List<Order> listOfOrders = new List<Order>();

            var rows = worksheet.RangeUsed().RowsUsed().ToList();
            for (int i = 1; i < rows.Count(); i++)
            {
                var row = rows[i];
                try
                {
                    if (productId == uint.Parse(row.Cell(2).Value.ToString()))
                    {
                        string sheetName = worksheet.Name;
                        uint id = uint.Parse(row.Cell(1).Value.ToString());
                        uint clientId = uint.Parse(row.Cell(3).Value.ToString());
                        uint numberOfOrder = uint.Parse(row.Cell(4).Value.ToString());
                        uint amount = uint.Parse(row.Cell(5).Value.ToString());
                        DateTime dateOfCreation = DateTime.Parse(row.Cell(6).Value.ToString());
                        listOfOrders.Add(new Order(id, sheetName, productId, clientId, numberOfOrder, amount, dateOfCreation));
                    }

                }
                catch (Exception ex)
                {
                    throw new Exception($"Unable to convert incoming strings to {nameof(Order)}", ex);
                }
            }
            return listOfOrders;
        }

        public List<Order> SelectForClientId(uint clientId)
        {
            List<Order> listOfOrders = new List<Order>();

            var rows = worksheet.RangeUsed().RowsUsed().ToList();
            for (int i = 1; i < rows.Count(); i++)
            {
                var row = rows[i];
                try
                {
                    if (clientId == uint.Parse(row.Cell(3).Value.ToString()))
                    {
                        string sheetName = worksheet.Name;
                        uint id = uint.Parse(row.Cell(1).Value.ToString());
                        uint productId = uint.Parse(row.Cell(2).Value.ToString());
                        uint numberOfOrder = uint.Parse(row.Cell(4).Value.ToString());
                        uint amount = uint.Parse(row.Cell(5).Value.ToString());
                        DateTime dateOfCreation = DateTime.Parse(row.Cell(6).Value.ToString());
                        listOfOrders.Add(new Order(id, sheetName, clientId, clientId, numberOfOrder, amount, dateOfCreation));
                    }

                }
                catch (Exception ex)
                {
                    throw new Exception($"Unable to convert incoming strings to {nameof(Order)}", ex);
                }
            }
            return listOfOrders;
        }
    }
}
