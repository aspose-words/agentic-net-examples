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

            // Create a placeholder image file.
            string imagePath = "sample.png";
            CreateSampleImage(imagePath);

            // Insert a floating image and configure its wrapping.
            Shape image = builder.InsertImage(imagePath);
            image.WrapType = WrapType.Square;
            image.RelativeHorizontalPosition = RelativeHorizontalPosition.Margin;
            image.RelativeVerticalPosition = RelativeVerticalPosition.Paragraph;
            image.AllowOverlap = true; // Allow the image to overlap other floating objects.

            // Move the builder to a new paragraph after the image.
            builder.Writeln();

            // Build a simple 2‑column table.
            Table table = builder.StartTable();
            builder.InsertCell();
            builder.Write("Cell 1");
            builder.InsertCell();
            builder.Write("Cell 2");
            builder.EndTable();

            // Configure the table to wrap text around it.
            table.PreferredWidth = PreferredWidth.FromPoints(300);
            table.TextWrapping = TextWrapping.Around;
            table.AbsoluteHorizontalDistance = 20; // Space between table and surrounding text.
            table.AbsoluteVerticalDistance = 20;   // Space between table and surrounding text.

            // Set the anchors so the table aligns with the image.
            table.HorizontalAnchor = RelativeHorizontalPosition.Margin;
            table.VerticalAnchor = RelativeVerticalPosition.Paragraph;

            // Save the document.
            string outputPath = "TableWrapAroundImage.docx";
            doc.Save(outputPath);

            // Verify that the output file was created.
            if (!File.Exists(outputPath))
                throw new FileNotFoundException("The output document was not saved.", outputPath);
        }

        // Helper method to create a placeholder PNG image.
        private static void CreateSampleImage(string path)
        {
            // A minimal 1x1 pixel PNG (transparent) encoded in base64.
            const string base64Png = "iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAQAAAC1HAwCAAAAC0lEQVR42mP8/x8AAwMCAO+XK6cAAAAASUVORK5CYII=";
            byte[] imageBytes = Convert.FromBase64String(base64Png);
            File.WriteAllBytes(path, imageBytes);
        }
    }
}
