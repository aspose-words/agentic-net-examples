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

        // ---------- Insert a floating image ----------
        // Tiny 1x1 PNG image encoded as Base64 (avoids System.Drawing dependency).
        byte[] pngBytes = Convert.FromBase64String(
            "iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAQAAAC1HAwCAAAAC0lEQVR42mP8/x8AAwMCAO+XK2cAAAAASUVORK5CYII=");
        Shape imageShape = builder.InsertImage(pngBytes);
        imageShape.WrapType = WrapType.Square;                     // Text wraps around the image.
        imageShape.RelativeHorizontalPosition = RelativeHorizontalPosition.Margin;
        imageShape.RelativeVerticalPosition = RelativeVerticalPosition.Paragraph;
        imageShape.HorizontalAlignment = HorizontalAlignment.Right; // Position the image to the right.
        imageShape.VerticalAlignment = VerticalAlignment.Top;
        imageShape.AllowOverlap = true;                            // Allow overlapping with other floating objects.

        // Add a paragraph after the image so the builder is positioned correctly.
        builder.Writeln();

        // ---------- Create a floating table that wraps text ----------
        Table table = builder.StartTable();

        // First row.
        builder.InsertCell();
        builder.Write("Cell 1");
        builder.InsertCell();
        builder.Write("Cell 2");
        builder.EndRow();

        // End the table construction.
        builder.EndTable();

        // Configure the table to float and wrap text around it.
        table.PreferredWidth = PreferredWidth.FromPoints(300);
        table.TextWrapping = TextWrapping.Around;          // Enable text wrapping.
        table.HorizontalAnchor = RelativeHorizontalPosition.Margin;
        table.VerticalAnchor = RelativeVerticalPosition.Paragraph;
        table.AbsoluteHorizontalDistance = 10;              // Space on the left/right of the table.
        table.AbsoluteVerticalDistance = 10;                // Space above/below the table.

        // Add some surrounding text to demonstrate wrapping.
        builder.Writeln("Lorem ipsum dolor sit amet, consectetur adipiscing elit, sed do eiusmod tempor incididunt ut labore et dolore magna aliqua.");

        // Save the document.
        string outputPath = Path.Combine(Directory.GetCurrentDirectory(), "TableWrapAroundImage.docx");
        doc.Save(outputPath);
    }
}
