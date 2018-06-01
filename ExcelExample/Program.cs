using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using System;
using System.Linq;

namespace ExcelExample
{
    public class Program
    {
        public static void Main(string[] args)
        {
            Console.WriteLine("Hello Excel World!");

            using (var spreadSheet = SpreadsheetDocument.Open("Names.xlsx", true))
            {
                var workbookPart = spreadSheet.WorkbookPart;
                var sstpart = workbookPart.GetPartsOfType<SharedStringTablePart>().First();
                var sharedStringTable = sstpart.SharedStringTable;

                var worksheetPart = workbookPart.WorksheetParts.First();
                var sheet = worksheetPart.Worksheet;

                var cells = sheet.Descendants<Cell>();
                var rows = sheet.Descendants<Row>();

                var enumerable = rows.ToList();

                Console.WriteLine($"Row count = {enumerable.LongCount()}");
                Console.WriteLine($"Cell count = {cells.LongCount()}");

                // Or... via each row
                foreach (var row in enumerable)
                {
                    foreach (var c in row.Elements<Cell>())
                    {
                        if (c.DataType != null && c.DataType == CellValues.SharedString)
                        {
                            var stringId = int.Parse(c.CellValue.Text);
                            var outstring = sharedStringTable.ChildElements[stringId].InnerText;
                            Console.Write($"Shared string {stringId}: {outstring} ");
                        }
                        else if (c.CellValue != null)
                        {
                            Console.Write($"Cell contents: { c.CellValue.Text}");
                        }
                    }
                    Console.WriteLine();
                }

                Console.ReadKey();
            }
        }
    }
}
