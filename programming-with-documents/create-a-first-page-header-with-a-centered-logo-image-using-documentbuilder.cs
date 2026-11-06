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

        // Enable a different header/footer for the first page.
        builder.PageSetup.DifferentFirstPageHeaderFooter = true;

        // Move the builder to the first page header.
        builder.MoveToHeaderFooter(HeaderFooterType.HeaderFirst);

        // Center the paragraph that will contain the logo.
        builder.ParagraphFormat.Alignment = ParagraphAlignment.Center;

        // A tiny placeholder PNG (1x1 transparent pixel) encoded in Base64.
        const string base64Png =
            "iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAQAAAC1HAwCAAAAC0lEQVR42mP8/x8AAwMCAO+XcZcAAAAASUVORK5CYII=";

        byte[] imageBytes = Convert.FromBase64String(base64Png);
        using (MemoryStream imageStream = new MemoryStream(imageBytes))
        {
            // Insert the image into the header.
            builder.InsertImage(imageStream);
        }

        // Return to the main document body.
        builder.MoveToDocumentEnd();

        // Add some sample body text.
        builder.Writeln("This is the document body.");

        // Save the document to disk.
        string outputPath = "FirstPageHeader.docx";
        doc.Save(outputPath);
        Console.WriteLine($"Document saved to {Path.GetFullPath(outputPath)}");
    }
}
