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

        // Base64‑encoded PNG image (1×1 pixel transparent).
        const string base64Image = "iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAQAAAC1HAwCAAAAC0lEQVR42mP8/x8AAwMCAO+X3V8AAAAASUVORK5CYII=";
        byte[] imageBytes = Convert.FromBase64String(base64Image);

        // Insert the image into the header.
        using (MemoryStream ms = new MemoryStream(imageBytes))
        {
            builder.InsertImage(ms);
        }

        // Save the document.
        const string outputPath = "FirstPageHeader.docx";
        doc.Save(outputPath);

        // Indicate completion.
        if (File.Exists(outputPath))
        {
            Console.WriteLine($"Document saved to: {Path.GetFullPath(outputPath)}");
        }
    }
}
