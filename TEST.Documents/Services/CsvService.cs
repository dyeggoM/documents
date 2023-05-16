using Microsoft.Extensions.Hosting;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace TEST.Documents.Services
{
    public class CsvService
    {
        private readonly IHostEnvironment _env;
        public CsvService(IHostEnvironment env)
        {
            _env = env;
        }

        public void CreateFile(dynamic data)
        {
            var dataString = "";
            foreach (var item in data)
            {
                dataString += $"{item.ColumnA},{item.ColumnB}\n";
            }
            var folderPath = Path.Combine(_env.ContentRootPath, "Files"); /* NOTA: _env.ContentRootPath is similar to: Directory.GetCurrentDirectory() */
            if (!Directory.Exists(folderPath))
                Directory.CreateDirectory(folderPath);
            var guid = Guid.NewGuid().ToString("N");
            string fileName = $"{guid}.csv";
            string filePath = Path.Combine(folderPath, fileName);
            using (var writer = new StreamWriter(System.IO.File.Open(filePath, FileMode.CreateNew), Encoding.UTF8))
            {
                writer.WriteLine("ColumnA,ColumnB");
                writer.WriteLine(dataString);
            }
        }

        public void ReadFile(string filePath)
        {
            using (var r = new StreamReader(filePath, Encoding.UTF8))
            {
                var data = r.ReadToEnd(); /*NOTA: JsonConvert.DeserializeObject<dynamic>(data) //If the object is in json format. */
                using (var reader = new StringReader(data))
                {
                    string line;
                    while ((line = reader.ReadLine()) != null)
                    {
                        var columnA = line.Split(',')[0];
                        var columnB = line.Split(',')[1];
                    }
                }
            }
        }
    }
}
