using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Drawing;
using Aspose.Words.Tables;
using System.Drawing;

namespace WatermarkTableExample
{
    public class Program
    {
        public static void Main()
        {
            // Create a new blank document.
            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);

            // -----------------------------------------------------------------
            // Build a simple 2x2 table.
            // -----------------------------------------------------------------
            builder.StartTable();

            // First row.
            builder.InsertCell();
            builder.Write("Cell 1, Row 1");
            builder.InsertCell();
            builder.Write("Cell 2, Row 1");
            builder.EndRow();

            // Second row.
            builder.InsertCell();
            builder.Write("Cell 1, Row 2");
            builder.InsertCell();
            builder.Write("Cell 2, Row 2");
            builder.EndRow();

            builder.EndTable();

            // -----------------------------------------------------------------
            // Add a text watermark that appears behind all content (including the table).
            // -----------------------------------------------------------------
            TextWatermarkOptions watermarkOptions = new TextWatermarkOptions
            {
                FontFamily = "Arial",
                FontSize = 72,                     // Large font size for visibility.
                Color = Color.LightGray,           // Light gray text color.
                Layout = WatermarkLayout.Diagonal, // Diagonal layout.
                IsSemitrasparent = false           // Fully opaque.
            };

            doc.Watermark.SetText("CONFIDENTIAL", watermarkOptions);

            // -----------------------------------------------------------------
            // Save the document.
            // -----------------------------------------------------------------
            string outputPath = Path.Combine(Directory.GetCurrentDirectory(), "WatermarkedTable.docx");
            doc.Save(outputPath);

            // Verify that the file was created.
            if (!File.Exists(outputPath))
                throw new Exception("The document was not saved successfully.");
        }
    }
}
