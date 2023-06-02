using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace TaskForDB.Entities
{
    public class Order : Entity
    {
        public uint ProductId { get; set; }
        public uint ClientId { get; set; }
        public uint NumberOfOrder { get; set; }
        public uint Amount { get; set; }
        public DateTime DateOfCreation { get; set; }

        public Order(uint id, string sheetName): base(id, sheetName) { }
        public Order(uint id, string sheetName, uint productId, uint clientId, uint numberOfOrder, uint amount, string dateOfCreation) : base(id, sheetName)
        {
            try
            {
                DateOfCreation = DateTime.Parse(dateOfCreation);
            }
            catch (Exception ex) 
            {
                throw new ArgumentException("Argument cannot be converted to DateTime", nameof(dateOfCreation), ex);
            }
            ProductId = productId;
            ClientId = clientId;
            NumberOfOrder = numberOfOrder;
            Amount = amount;
        }
        public Order(uint id, string sheetName, uint productId, uint clientId, uint numberOfOrder, uint amount, DateTime dateOfCreation) : base(id, sheetName)
        {
            ProductId = productId;
            ClientId = clientId;
            NumberOfOrder = numberOfOrder;
            Amount = amount;
            DateOfCreation = dateOfCreation;
        }
    }
}
