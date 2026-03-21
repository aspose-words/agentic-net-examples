using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Tables;
using Aspose.Words.Drawing;

class TableWrapAroundImageExample
{
    static void Main()
    {
        // Create a new blank document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Create a tiny PNG image (1x1 pixel, white) in memory.
        byte[] pngBytes = Convert.FromBase64String(
            "iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAQAAAC1HAwCAAAAC0lEQVR42mP8/x8AAwMCAO+XK2cAAAAASUVORK5CYII=");
        using var imageStream = new MemoryStream(pngBytes);

        // Insert a floating image that will be wrapped by surrounding text.
        Shape image = builder.InsertImage(imageStream);
        image.WrapType = WrapType.Square;                     // Text wraps on all sides of the image.
        image.RelativeHorizontalPosition = RelativeHorizontalPosition.Margin;
        image.RelativeVerticalPosition = RelativeVerticalPosition.Paragraph;
        image.HorizontalAlignment = HorizontalAlignment.Left; // Align image to the left side.
        image.VerticalAlignment = VerticalAlignment.Top;
        image.BehindText = false;                            // Image appears in front of text.

        // Add a paragraph after the image to contain the table.
        builder.Writeln();

        // Start building a table that will wrap around the floating image.
        Table table = builder.StartTable();
        builder.InsertCell();
        builder.Write("Cell 1");
        builder.InsertCell();
        builder.Write("Cell 2");
        builder.EndRow();
        builder.EndTable();

        // Enable text wrapping for the table.
        table.TextWrapping = TextWrapping.Around;            // Wrap text around the table.

        // Adjust the distance between the table and surrounding text.
        table.AbsoluteHorizontalDistance = 10; // Points of horizontal space.
        table.AbsoluteVerticalDistance = 10;   // Points of vertical space.

        // Optionally set the anchor positions to control where the table floats.
        table.HorizontalAnchor = RelativeHorizontalPosition.Margin;
        table.VerticalAnchor = RelativeVerticalPosition.Paragraph;

        // Ensure output directory exists.
        string outputDir = Path.Combine(Directory.GetCurrentDirectory(), "Output");
        Directory.CreateDirectory(outputDir);

        // Save the resulting document.
        string outputPath = Path.Combine(outputDir, "TableWrapAroundImage.docx");
        doc.Save(outputPath);
        Console.WriteLine($"Document saved to: {outputPath}");
    }
}
