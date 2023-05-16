using System;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Threading.Tasks;
using iTextSharp.text;
using iTextSharp.text.pdf;
using Microsoft.Extensions.Hosting;

namespace TEST.Documents.Services
{
    public class PdfService
    {
        private readonly IHostEnvironment _env;
        public PdfService(IHostEnvironment env)
        {
            _env = env;
        }

        public void CreateFile(dynamic data)
        {
            var today = DateTime.Now;
            var guid = Guid.NewGuid().ToString("N");
            var fileName = $"{guid}-TEMP.pdf";
            var folderPath = Path.Combine(_env.ContentRootPath, "Files"); /* NOTA: _env.ContentRootPath is similar to: Directory.GetCurrentDirectory() */
            if (!Directory.Exists(folderPath))
                Directory.CreateDirectory(folderPath);
            var filePath = Path.Combine(folderPath, fileName);
            var fileStream = new FileStream(filePath, FileMode.Create);
            var pdfDoc = new Document(PageSize.A4);
            var pdfWriter = PdfWriter.GetInstance(pdfDoc, fileStream);
            pdfDoc.AddAuthor("Author");
            pdfDoc.SetMargins(36f, 36f, 36f, 36f);
            
            #region FontDefinitions
            var fontName = "Arial";
            var colorGray = new BaseColor(0, 111, 183);
            var defaultFont = new Font(FontFactory.GetFont(fontName).BaseFont, 8);
            var tableFont = new Font(FontFactory.GetFont(fontName).BaseFont, 6, Font.NORMAL, colorGray);
            var tableTitleFont = new Font(FontFactory.GetFont(fontName).BaseFont, 6, Font.BOLD);
            #endregion

            pdfDoc.Open();

            #region Intro
            var ci = new CultureInfo("es-CO");
            var dateMessage = $"Medellín {today.Date.ToString("dd", ci)} de {today.Date.ToString("MMMM", ci)} de {today.Date.ToString("yyyy", ci)}";
            Paragraph fecha = new Paragraph(dateMessage, defaultFont);
            fecha.Alignment = Element.ALIGN_JUSTIFIED;
            fecha.SpacingAfter = 24f;
            pdfDoc.Add(fecha);
            if (!String.IsNullOrWhiteSpace(data.NombreCliente) && !String.IsNullOrWhiteSpace(data.Encabezado))
            {
                Paragraph encabezado = new Paragraph(data.Encabezado + "\n" + data.NombreCliente, defaultFont);
                encabezado.Alignment = Element.ALIGN_JUSTIFIED;
                encabezado.SpacingAfter = 12f;
                pdfDoc.Add(encabezado);
            }
            Paragraph saludo = new Paragraph("Cordial saludo", defaultFont);
            saludo.Alignment = Element.ALIGN_JUSTIFIED;
            saludo.SpacingAfter = 12f;
            pdfDoc.Add(saludo);
            Paragraph mensaje = new Paragraph("Por medio de la presente damos respuesta a la solicitud ...", defaultFont);
            mensaje.Alignment = Element.ALIGN_JUSTIFIED;
            mensaje.SpacingAfter = 12f;
            pdfDoc.Add(mensaje);
            #endregion

            #region Tabla
            var i = 0;
            var table = new PdfPTable(3);
            Rectangle rect = new Rectangle(523, 770);
            table.SetWidthPercentage(new float[] { 100, 350, 50 }, rect);
            table.AddCell(new PdfPCell(new Phrase(new Chunk("Columna A", tableTitleFont))) { Border = Rectangle.BOTTOM_BORDER });
            table.AddCell(new PdfPCell(new Phrase(new Chunk("Columna B", tableTitleFont))) { Border = Rectangle.BOTTOM_BORDER });
            table.AddCell(new PdfPCell(new Phrase(new Chunk("Columna C", tableTitleFont))) { Border = Rectangle.BOTTOM_BORDER, HorizontalAlignment = Element.ALIGN_CENTER });
            foreach (var item in data.ItemList)
            {
                var backColor = (i % 2 == 0) ? BaseColor.LightGray : BaseColor.White;
                table.AddCell(new PdfPCell(new Phrase(new Chunk(item.ColumnA, tableFont))) { Border = Rectangle.RIGHT_BORDER, BackgroundColor = backColor });
                table.AddCell(new PdfPCell(new Phrase(new Chunk(item.ColumnB, tableFont))) { Border = Rectangle.NO_BORDER, BackgroundColor = backColor });
                table.AddCell(new PdfPCell(new Phrase(new Chunk(item.ColumnC, tableFont))) { Border = Rectangle.NO_BORDER, HorizontalAlignment = Element.ALIGN_CENTER, BackgroundColor = backColor });
                i++;
            }
            pdfDoc.Add(table);
            #endregion
            pdfDoc.Close();
            fileStream.Dispose();
        }

        public void AddWatermark(string noWatermarkFilePath, string folderPath)
        {
            var marca = Path.Combine(_env.ContentRootPath, Path.Combine("Resources", "MarcaAgua.png"));
            var watermark = Image.GetInstance(marca);
            watermark.ScaleAbsolute(PageSize.A4.Width, PageSize.A4.Height);
            watermark.SetAbsolutePosition(0f, 0f);
            var pdfReader = new PdfReader(noWatermarkFilePath);
            var guid = Guid.NewGuid().ToString("N");
            var fileName = $"{guid}.pdf";
            var filePath = Path.Combine(folderPath, fileName);
            var outStream = new FileStream(filePath, FileMode.Create);
            var pdfStamper = new PdfStamper(pdfReader, outStream);
            int total = pdfReader.NumberOfPages + 1;
            var pageSize = pdfReader.GetPageSize(1);
            var width = pageSize.Width;
            var height = pageSize.Height;
            var font = BaseFont.CreateFont();
            var gs = new PdfGState();
            for (var i = 1; i < total; i++)
            {
                string pagination = $"Página {i} de {total - 1}";
                var content = pdfStamper.GetOverContent(i);
                gs.FillOpacity = 0.2f;
                content.SetGState(gs);
                content.BeginText();
                content.AddImage(watermark);
                content.SetColorFill(BaseColor.Black);
                content.SetFontAndSize(font, 5);
                content.SetTextMatrix(0, 0);
                content.ShowTextAligned(
                    Element.ALIGN_RIGHT,
                    pagination,
                    width - 36,
                    20,
                    0);
                content.EndText();
            }
            pdfStamper.Close();
            pdfReader.Close();
            outStream.Dispose();
        }
    }
}
