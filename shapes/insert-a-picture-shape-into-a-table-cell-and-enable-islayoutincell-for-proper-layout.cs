using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Drawing;

public class Program
{
    public static void Main()
    {
        // Create a new blank document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Build a simple 2‑column table.
        builder.StartTable();
        builder.InsertCell(); // First cell – will contain the image.
        builder.InsertCell(); // Second cell – placeholder content.
        builder.Write("Second cell");
        builder.EndRow();
        builder.EndTable();

        // Move the cursor to the first cell where the image will be placed.
        builder.MoveTo(doc.FirstSection.Body.Tables[0].FirstRow.FirstCell.FirstParagraph);

        // A tiny 1×1 PNG image (transparent) encoded in Base64.
        // This avoids the need for System.Drawing dependencies.
        string base64Png = "iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAQAAAC1HAwCAAAAC0lEQVR42mP8/x8AAwMCAO+X2V8AAAAASUVORK5CYII=";
        byte[] imageBytes = Convert.FromBase64String(base64Png);

        // Insert the image as a floating shape with explicit size and no text wrapping.
        Shape picture = builder.InsertImage(
            imageBytes,
            RelativeHorizontalPosition.Margin, 0,   // left offset
            RelativeVerticalPosition.Margin, 0,     // top offset
            100, 100,                               // width & height in points
            WrapType.None);                         // no text wrapping

        // Enable layout inside the table cell.
        picture.IsLayoutInCell = true;

        // Save the document to the current directory.
        string outputPath = Path.Combine(Directory.GetCurrentDirectory(), "TableImage.docx");
        doc.Save(outputPath);

        // Validate that the file was created.
        if (!File.Exists(outputPath))
            throw new InvalidOperationException("The document was not saved correctly.");
    }
}
