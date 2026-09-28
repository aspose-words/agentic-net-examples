using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Drawing;

public class Program
{
    public static void Main()
    {
        // Create a new document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Enable a different header/footer for the first page.
        builder.PageSetup.DifferentFirstPageHeaderFooter = true;

        // Move to the first page header.
        builder.MoveToHeaderFooter(HeaderFooterType.HeaderFirst);

        // Center the content in the header.
        builder.ParagraphFormat.Alignment = ParagraphAlignment.Center;

        // Base64-encoded PNG (1x1 red pixel) to use as a logo.
        string base64Image = "iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAQAAAC1HAwCAAAAC0lEQVR42mP8/x8AAwMCAO+XK6cAAAAASUVORK5CYII=";
        byte[] imageBytes = Convert.FromBase64String(base64Image);

        // Insert the image into the header.
        using (MemoryStream imageStream = new MemoryStream(imageBytes))
        {
            builder.InsertImage(imageStream);
        }

        // Save the document.
        string outputPath = "FirstPageHeaderWithLogo.docx";
        doc.Save(outputPath);

        // Indicate completion (optional).
        Console.WriteLine($"Document saved to: {Path.GetFullPath(outputPath)}");
    }
}
