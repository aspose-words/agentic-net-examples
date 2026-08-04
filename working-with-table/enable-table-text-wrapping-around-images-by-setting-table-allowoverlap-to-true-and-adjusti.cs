using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Drawing;
using Aspose.Words.Tables;

public class Program
{
    public static void Main()
    {
        // Create a new blank document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // -----------------------------------------------------------------
        // Insert a floating image that will allow text to wrap around it.
        // -----------------------------------------------------------------
        const string imagePath = "sample.png";
        if (!File.Exists(imagePath))
        {
            // Create a minimal 1x1 transparent PNG if it does not exist.
            const string base64Png = "iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAQAAAC1HAwCAAAAC0lEQVR42mP8/x8AAwMCAO+X6V8AAAAASUVORK5CYII=";
            byte[] pngBytes = Convert.FromBase64String(base64Png);
            File.WriteAllBytes(imagePath, pngBytes);
        }

        // Insert the image as a floating shape.
        Shape imageShape = builder.InsertImage(imagePath);
        imageShape.WrapType = WrapType.Square;                     // Wrap text tightly around the image.
        imageShape.RelativeHorizontalPosition = RelativeHorizontalPosition.Margin;
        imageShape.RelativeVerticalPosition = RelativeVerticalPosition.Paragraph;
        imageShape.AllowOverlap = true;                            // Allow the image to overlap other floating objects.

        // Add a paragraph of text before the table.
        builder.Writeln("This paragraph appears before the table. The image above should have text wrapped around it.");

        // -----------------------------------------------------------------
        // Create a floating table that will wrap text around it.
        // -----------------------------------------------------------------
        Table table = builder.StartTable();
        builder.InsertCell();
        builder.Write("Cell 1");
        builder.InsertCell();
        builder.Write("Cell 2");
        builder.EndRow();
        builder.EndTable();

        // Enable text wrapping around the table and make it a floating object.
        table.TextWrapping = TextWrapping.Around;
        table.HorizontalAnchor = RelativeHorizontalPosition.Margin;   // Position relative to page margin horizontally.
        table.VerticalAnchor = RelativeVerticalPosition.Paragraph;   // Position relative to the paragraph vertically.
        table.AbsoluteHorizontalDistance = 20;                        // Horizontal offset from the anchor point.
        table.AbsoluteVerticalDistance = 20;                          // Vertical offset from the anchor point.

        // Note: Table.AllowOverlap is read‑only and defaults to true. No explicit check is required.

        // Add another paragraph after the table.
        builder.Writeln("This paragraph appears after the table. Both the image and the table should have text wrapped around them.");

        // -----------------------------------------------------------------
        // Save the document.
        // -----------------------------------------------------------------
        const string outputDir = "Output";
        Directory.CreateDirectory(outputDir);
        string outputPath = Path.Combine(outputDir, "TableWrapAroundImage.docx");
        doc.Save(outputPath);
    }
}
