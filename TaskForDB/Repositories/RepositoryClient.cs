using ClosedXML.Excel;
using DocumentFormat.OpenXml.Office2010.Excel;
using DocumentFormat.OpenXml.Spreadsheet;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using TaskForDB.Entities;

namespace TaskForDB.Repositories
{
    internal class RepositoryClient : Repository<Client>
    {
        public RepositoryClient(IXLWorksheet worksheet) : base(worksheet) { }

        public override List<Client> SelectAll()
        {
            List<Client> listOfClients = new List<Client>();

            var rows = worksheet.RangeUsed().RowsUsed().ToList();
            for (int i = 1; i < rows.Count(); i++) 
            {
                var row = rows[i];
                try
                {
                    string sheetName = worksheet.Name;
                    uint id = uint.Parse(row.Cell(1).Value.ToString());
                    string organizationName = row.Cell(2).Value.ToString();
                    string address = row.Cell(3).Value.ToString();
                    string contactName = row.Cell(4).Value.ToString();
                    Client client = new Client(id, sheetName, organizationName, address, contactName);
                    listOfClients.Add(client);

                }
                catch (Exception ex) 
                {
                    throw new Exception($"Unable to convert incoming strings to {nameof(Client)}", ex);
                }
            }

            return listOfClients;
        }
        public override Client? GetForId(uint id)
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
                        string organizationName = row.Cell(2).Value.ToString();
                        string address = row.Cell(3).Value.ToString();
                        string contactName = row.Cell(4).Value.ToString();
                        return new Client(id, sheetName, organizationName, address, contactName);
                    }
                }
                catch (Exception ex)
                {
                    throw new Exception($"Unable to convert incoming strings to {nameof(Client)}", ex);
                }
            }
            return null;
        }
        public Client? GetForOrganizationName(string organizationName)
        {
            var rows = worksheet.RangeUsed().RowsUsed().ToList();
            for (int i = 1; i < rows.Count(); i++)
            {
                var row = rows[i];
                try
                {
                    if (organizationName == row.Cell(2).Value.ToString())
                    {
                        string sheetName = worksheet.Name;
                        uint id = uint.Parse(row.Cell(1).Value.ToString());
                        string address = row.Cell(3).Value.ToString();
                        string contactName = row.Cell(4).Value.ToString();
                        return new Client(id, sheetName, organizationName, address, contactName);
                    }
                }
                catch (Exception ex)
                {
                    throw new Exception($"Unable to convert incoming strings to {nameof(Client)}", ex);
                }
            }
            return null;
        }
        public IXLWorksheet ChangeContactName(uint id, string newContactName)
        {
            var rows = worksheet.RangeUsed().RowsUsed().ToList();
            for (int i = 1; i < rows.Count(); i++)
            {
                var row = rows[i];
                if (id == uint.Parse(row.Cell(1).Value.ToString()))
                {
                    worksheet.Row(row.RowNumber()).Cell(4).Value = newContactName;
                }
            }
            return worksheet;
        }
    }
}
