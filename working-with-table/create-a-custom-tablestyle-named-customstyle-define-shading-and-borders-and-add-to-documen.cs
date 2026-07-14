using System;
using System.Drawing;
using Aspose.Words;
using Aspose.Words.Tables;

namespace TableStyleExample
{
    public class Program
    {
        public static void Main()
        {
            // Create a new blank document.
            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);

            // Build a simple 2x2 table.
            Table table = builder.StartTable();

            // First row.
            builder.InsertCell();
            builder.Write("Header 1");
            builder.InsertCell();
            builder.Write("Header 2");
            builder.EndRow();

            // Second row.
            builder.InsertCell();
            builder.Write("Cell 1");
            builder.InsertCell();
            builder.Write("Cell 2");
            builder.EndTable();

            // Create a custom table style named "CustomStyle".
            TableStyle customStyle = (TableStyle)doc.Styles.Add(StyleType.Table, "CustomStyle");

            // Define shading (background color) for the style.
            customStyle.Shading.BackgroundPatternColor = Color.LightYellow;

            // Define borders for the style.
            customStyle.Borders.Color = Color.DarkBlue;
            customStyle.Borders.LineStyle = LineStyle.Single;
            customStyle.Borders.LineWidth = 1.5; // Optional: set line width.

            // Apply the custom style to the table.
            table.Style = customStyle;

            // Save the document to the local file system.
            string outputPath = "CustomTableStyle.docx";
            doc.Save(outputPath);

            // Simple verification that the file was created.
            if (System.IO.File.Exists(outputPath))
            {
                Console.WriteLine($"Document saved successfully to '{outputPath}'.");
            }
            else
            {
                throw new InvalidOperationException("Failed to save the document.");
            }
        }
    }
}
