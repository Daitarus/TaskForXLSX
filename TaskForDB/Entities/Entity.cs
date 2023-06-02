using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace TaskForDB.Entities
{
    public abstract class Entity
    {
        public readonly string sheetName;

        public uint Id { get; set; }

        public Entity(uint id, string sheetName)
        {
            Id = id;
            this.sheetName = sheetName;
        }
    }
}
