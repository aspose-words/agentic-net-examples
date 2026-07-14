using System;
using System.IO;
using System.Drawing;
using Aspose.Words;
using Aspose.Words.Drawing;
using Aspose.Words.Tables;

namespace WatermarkBehindTableExample
{
    public class Program
    {
        public static void Main()
        {
            // Create a new blank document.
            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);

            // Add a text watermark that will appear behind all document content,
            // including the table that will be added later.
            TextWatermarkOptions watermarkOptions = new TextWatermarkOptions
            {
                FontFamily = "Arial",
                FontSize = 36,
                Color = Color.Gray,
                Layout = WatermarkLayout.Diagonal,
                IsSemitrasparent = false
            };
            doc.Watermark.SetText("CONFIDENTIAL", watermarkOptions);

            // Build a simple 2x2 table that will appear above the watermark.
            Table table = builder.StartTable();

            // First row.
            builder.InsertCell();
            builder.Write("Cell 1");
            builder.InsertCell();
            builder.Write("Cell 2");
            builder.EndRow();

            // Second row.
            builder.InsertCell();
            builder.Write("Cell 3");
            builder.InsertCell();
            builder.Write("Cell 4");
            builder.EndRow();

            // Finish the table.
            builder.EndTable();

            // Save the document to the current directory.
            string outputPath = Path.Combine(Directory.GetCurrentDirectory(), "WatermarkTable.docx");
            doc.Save(outputPath);
        }
    }
}
