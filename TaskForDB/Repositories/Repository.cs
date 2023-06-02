using ClosedXML.Excel;
using DocumentFormat.OpenXml.Spreadsheet;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace TaskForDB.Repositories
{
    public abstract class Repository<T> where T : Entities.Entity 
    {
        protected readonly IXLWorksheet worksheet;
        public Repository(IXLWorksheet worksheet) 
        { 
            this.worksheet = worksheet;
        }

        public abstract List<T> SelectAll();
        public abstract T? GetForId(uint id);
    }
}
