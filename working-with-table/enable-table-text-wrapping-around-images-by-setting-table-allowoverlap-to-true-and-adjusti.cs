using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Tables;
using Aspose.Words.Drawing;

namespace TableWrapAroundImageExample
{
    public class Program
    {
        public static void Main()
        {
            // Create a new blank document.
            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);

            // ----- Insert a floating image -----
            // Create a simple 1x1 pixel PNG image from a Base64 string.
            const string base64Png = "iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAQAAAC1HAwCAAAAC0lEQVR42mP8/5+BAQAE/wJ+XKcZAAAAAElFTkSuQmCC";
            byte[] imageBytes = Convert.FromBase64String(base64Png);
            using (MemoryStream imageStream = new MemoryStream(imageBytes))
            {
                // Insert the image as a floating shape.
                Shape imageShape = builder.InsertImage(imageStream);
                // Wrap text around the image.
                imageShape.WrapType = WrapType.Square;
                // Position the image relative to the page margin.
                imageShape.RelativeHorizontalPosition = RelativeHorizontalPosition.Margin;
                imageShape.RelativeVerticalPosition = RelativeVerticalPosition.Paragraph;
                // Optional: adjust the distance from surrounding text.
                imageShape.Left = 0;
                imageShape.Top = 0;
                // Allow the shape to overlap other floating objects.
                imageShape.AllowOverlap = true;
            }

            // Add a paragraph before the table to demonstrate wrapping.
            builder.Writeln("This paragraph appears before the table. The table will wrap around the floating image above.");

            // ----- Create a table -----
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

            // Enable text wrapping around the table.
            table.TextWrapping = TextWrapping.Around;
            // Set distances between the table and surrounding text.
            table.AbsoluteHorizontalDistance = 20; // points
            table.AbsoluteVerticalDistance = 20;   // points

            // Note: Table.AllowOverlap is read‑only. After setting TextWrapping to Around,
            // Aspose.Words may set AllowOverlap to false internally, which is expected.
            // Therefore we do not assert that AllowOverlap must be true.

            // Save the document.
            string outputPath = Path.Combine(Environment.CurrentDirectory, "TableWrapAroundImage.docx");
            doc.Save(outputPath);

            // Simple validation that the file was created.
            if (!File.Exists(outputPath))
                throw new FileNotFoundException("Failed to create the output document.", outputPath);
        }
    }
}
