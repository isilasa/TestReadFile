using Newtonsoft.Json;
using System;
using System.IO;
using TestReadFile.Interfaces;
using TestReadFile.Models;

namespace TestReadFile.Writers
{
    public class JsonWrite : IWriter
    {
        public void Write(Sheet1[] rows, string path)
        {
            try
            {
                JsonSerializer serializer = new JsonSerializer();
                var file = $"{path}/SaveFile.json";
                using (StreamWriter sw = new StreamWriter(file))
                using (JsonWriter writer = new JsonTextWriter(sw))
                {
                    foreach (var row in rows)
                    {
                        serializer.Serialize(writer, row);
                    }
                }
                Console.WriteLine($"Fale save in {file}");
            }
            catch (Exception ex)
            {
                Console.WriteLine(ex.Message);
            }
        }
    }
}
