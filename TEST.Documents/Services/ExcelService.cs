using System;
using ClosedXML.Excel;
using System.Collections.Generic;
using System.Linq;
using System.Threading.Tasks;
using System.IO;
using Microsoft.Extensions.Hosting;

namespace TEST.Documents.Services
{
    public class ExcelService
    {
        private readonly IHostEnvironment _env;
        public ExcelService(IHostEnvironment env)
        {
            _env = env;
        }

        public void CreateFile(dynamic data)
        {
            var workbook = new XLWorkbook();
            var ws = workbook.Worksheets.Add("Hoja 1");
            ws.Cell("A1").Value = "Columna A";
            ws.Cell("B1").Value = "Columna B";
            ws.Cell("C1").Value = "Columna C";
            var headersRange = ws.Range("A1:C1");
            headersRange.Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Center;
            headersRange.Style.Font.Bold = true;
            int rowIndex = 2;
            foreach (var item in data)
            {
                ws.Cell($"A{rowIndex}").SetValue<string>(item.ColumnA);
                ws.Cell($"B{rowIndex}").SetValue<string>(item.ColumnB);
                ws.Cell($"C{rowIndex}").SetValue<string>(item.ColumnC);
                rowIndex++;
            }
            ws.Columns(1, 3).AdjustToContents();
            var guid = Guid.NewGuid().ToString("N");
            var fileName = $"{guid}.xlsx";
            var folderPath = Path.Combine(_env.ContentRootPath, "Files"); /* NOTA: _env.ContentRootPath is similar to: Directory.GetCurrentDirectory() */
            if (!Directory.Exists(folderPath))
                Directory.CreateDirectory(folderPath);
            var excelFilePath = Path.Combine(folderPath, fileName);
            workbook.SaveAs(excelFilePath);
        }

        public void ReadFile(string filePath)
        {
            var workbook = new XLWorkbook(filePath);
            var workSheets = workbook.Worksheets;
            foreach (var ws in workSheets)
            {
                var range = ws.RangeUsed();
                for (var i = 2; i <= range.RowCount(); i++)
                {
                    var columnA = range.Cell(i, "A").Value.ToString();
                    var columnB = range.Cell(i, "B").Value.ToString();
                    var columnC = range.Cell(i, "C").Value.ToString();
                }
            }
        }
    }
}
