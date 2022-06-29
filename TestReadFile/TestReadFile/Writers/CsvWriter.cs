using System;
using System.Globalization;
using System.IO;
using System.Text;
using TestReadFile.Interfaces;
using TestReadFile.Models;

namespace TestReadFile.Writers
{
    public class CsvWriter : IWriter
    {
        public void Write(Sheet1[] rows, string path)
        {
            try
            {
                var file = $"{path}/SaveFile.csv";
                var csv = new StringBuilder();
                var newLine =
                        "Id;" +
                        "Segment;" +
                        "Country;" +
                        "Product;" +
                        "DiscountBand;" +
                        "UnitsSold;" +
                        "ManufacturingPrice;" +
                        "SalePrice;" +
                        "GrossSales;" +
                        "Discounts;" +
                        "Sales;" +
                        "COGS;" +
                        "Profit;" +
                        "Date;" +
                        "MonthNumber;" +
                        "MonthName;" +
                        "Year;";
                csv.AppendLine(newLine);
                foreach (var row in rows)
                {
                    newLine =
                        $"{row.Id};" +
                        $"{row.Segment};" +
                        $"{row.Country};" +
                        $"{row.Product};" +
                        $"{row.DiscountBand};" +
                        $"{row.UnitsSold};" +
                        $"{row.ManufacturingPrice};" +
                        $"{row.SalePrice};" +
                        $"{row.GrossSales};" +
                        $"{row.Discounts};" +
                        $"{row.Sales};" +
                        $"{row.COGS};" +
                        $"{row.Profit};" +
                        $"{row.Date:dd.MM.yyyy};" +
                        $"{row.Date.Month};" +
                        $"{row.Date.ToString("MMMM", CultureInfo.GetCultureInfo("en-En").DateTimeFormat)};" +
                        $"{row.Year};";
                    csv.AppendLine(newLine);
                }
                File.WriteAllText(file, csv.ToString());
                Console.WriteLine($"Fale save in {file}");
            }
            catch (Exception ex)
            {
                Console.WriteLine(ex.Message);
            }
        }
    }
}
