using LinqToExcel;
using System;
using System.IO;
using System.Linq;
using TestReadFile.Models;
using TestReadFile.Writers;

namespace TestReadFile
{
    public class Program
    {
        static void Main(string[] args)
        {
            Console.WriteLine("Input path to file");
            Console.Write("Path to excel file: ");

            var pathToExcelFile = Console.ReadLine();
            while (!File.Exists(pathToExcelFile))
            {
                Console.WriteLine("File does not exist. Input path again.");
                Console.Write("Path to excel file: ");
                pathToExcelFile = Console.ReadLine();
            }

            var excelFile = new ExcelQueryFactory(pathToExcelFile);
            try
            {
                var rows = excelFile.Worksheet<Sheet1>()
                    .Where(profit => profit.Profit > 100000).ToArray();

                Console.WriteLine("Choose which format to save.");
                Console.WriteLine(".csv input 1, .json input 2");
                Console.Write("Your choose: ");
                var key = Console.ReadKey();

                string savePath;
                bool isNeedChoose = true;
                while (isNeedChoose)
                {
                    Int32.TryParse(key.KeyChar.ToString(), out var inputKey);
                    switch (inputKey)
                    {
                        case 1:
                            Console.WriteLine("\nInput path where need save file");
                            Console.Write("Path : ");
                            savePath = Console.ReadLine();

                            var csvWriter = new CsvWriter();
                            csvWriter.Write(rows,savePath);

                            isNeedChoose = false;
                            return;
                        case 2:
                            Console.WriteLine("\nInput path where need save file");
                            Console.Write("Path : ");
                            savePath = Console.ReadLine();

                            var writer = new JsonWrite();
                            writer.Write(rows,savePath);

                            isNeedChoose = false;
                            return;
                        default:
                            Console.WriteLine("\nFormat does not exist. Choose format again.");
                            Console.WriteLine(".csv input 1, .json input 2");
                            Console.Write("Your choose: ");
                            key = Console.ReadKey();
                            break;
                    }
                }

                Console.WriteLine("\nЗапись завершена");
                Console.ReadKey();
            }
            catch (Exception ex)
            {
                Console.WriteLine(ex.Message);
            }
        }
    }
}