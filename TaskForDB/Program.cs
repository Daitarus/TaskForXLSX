using ClosedXML.Excel;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Spreadsheet;
using TaskForDB.Entities;
using TaskForDB.Repositories;

namespace TaskForDB
{
    internal class Program
    {
        static void Main(string[] args)
        {
            ConsoleLayout consoleLayout = new ConsoleLayout();
            consoleLayout.MainLayout();
        }
    }
}