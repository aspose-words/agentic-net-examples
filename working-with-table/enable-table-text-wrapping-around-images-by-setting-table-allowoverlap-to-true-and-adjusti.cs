using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Tables;
using Aspose.Words.Drawing;

public class Program
{
    public static void Main()
    {
        // Create a new blank document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // -----------------------------------------------------------------
        // Insert a floating image (a 1x1 pixel PNG generated from a base‑64 string).
        // -----------------------------------------------------------------
        byte[] pngBytes = Convert.FromBase64String(
            "iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAQAAAC1HAwCAAAAC0lEQVR42mP8/x8AAwMCAO+XK6cAAAAASUVORK5CYII=");

        // Insert the image as a floating shape with the desired size and position.
        Shape image = builder.InsertImage(
            pngBytes,
            RelativeHorizontalPosition.Page, 50,   // left distance from page
            RelativeVerticalPosition.Page, 50,     // top distance from page
            100,                                   // width in points
            100,                                   // height in points
            WrapType.Square);                     // text will wrap around the image

        // Add a paragraph of text before the table – this text will flow around the image.
        builder.Writeln("Lorem ipsum dolor sit amet, consectetur adipiscing elit. " +
                        "Sed do eiusmod tempor incididunt ut labore et dolore magna aliqua.");

        // -----------------------------------------------------------------
        // Build a simple 2×2 table.
        // -----------------------------------------------------------------
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
        table = builder.EndTable();

        // -----------------------------------------------------------------
        // Enable text wrapping around the table and adjust its layout.
        // -----------------------------------------------------------------
        // Verify that the table allows overlap (default is true).
        if (!table.AllowOverlap)
            throw new InvalidOperationException("Table does not allow overlap, but it should.");

        // Wrap text around the table.
        table.TextWrapping = TextWrapping.Around;

        // Position the floating table relative to the page margins.
        table.HorizontalAnchor = RelativeHorizontalPosition.Margin;
        table.VerticalAnchor = RelativeVerticalPosition.Paragraph;

        // Set distances between the table and surrounding text (in points).
        table.AbsoluteHorizontalDistance = 10;
        table.AbsoluteVerticalDistance = 10;

        // Add more text after the table to demonstrate wrapping.
        builder.Writeln("Additional paragraph that follows the table. " +
                        "The text should wrap around both the image and the table according to the layout settings.");

        // -----------------------------------------------------------------
        // Save the document to the current working directory.
        // -----------------------------------------------------------------
        string outputPath = Path.Combine(Directory.GetCurrentDirectory(), "TableWrapAround.docx");
        doc.Save(outputPath);

        // Simple validation that the file was created.
        if (!File.Exists(outputPath))
            throw new FileNotFoundException("The output document was not created.", outputPath);
    }
}
