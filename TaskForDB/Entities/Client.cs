using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace TaskForDB.Entities
{
    internal class Client : Entity
    {
        public string OrganizationName { get; set; }
        public string Address { get; set; }
        public string ContactName { get; set; }

        public Client(uint id, string sheetName) :base(id, sheetName) { }
        public Client(uint id, string sheetName, string organizationName, string address, string contactName) : base(id, sheetName)
        {
            OrganizationName = organizationName;
            Address = address;
            ContactName = contactName;
        }
    }
}
