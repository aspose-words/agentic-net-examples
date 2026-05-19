using System;
using Aspose.Words;
using Aspose.Words.Drawing;

public class Program
{
    public static void Main()
    {
        // Create a new blank document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Enable a different header for the first page.
        builder.PageSetup.DifferentFirstPageHeaderFooter = true;

        // Move the cursor to the first‑page header.
        builder.MoveToHeaderFooter(HeaderFooterType.HeaderFirst);

        // A tiny PNG image (1x1 pixel, transparent) encoded in Base64.
        // This avoids the need for System.Drawing and external image files.
        const string base64Png = "iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAQAAAC1HAwCAAAAC0lEQVR42mP8/5+BAQAE/wJ/lK5XAAAAAElFTkSuQmCC";
        byte[] imageBytes = Convert.FromBase64String(base64Png);

        // Insert the image into the header. The method returns the Shape that represents the image.
        Shape shape = builder.InsertImage(imageBytes);

        // Make the image float (not inline) and place it at the center of the page.
        shape.WrapType = WrapType.None;
        shape.BehindText = true;
        shape.RelativeHorizontalPosition = RelativeHorizontalPosition.Page;
        shape.RelativeVerticalPosition = RelativeVerticalPosition.Page;
        shape.HorizontalAlignment = HorizontalAlignment.Center;
        shape.VerticalAlignment = VerticalAlignment.Center;

        // Add some content to the document body to demonstrate the header.
        builder.MoveToSection(0);
        builder.Writeln("First page with centered logo in the header.");
        builder.InsertBreak(BreakType.PageBreak);
        builder.Writeln("Second page – the first‑page header does not appear here.");

        // Save the document to the local file system.
        string outputPath = "FirstPageHeader.docx";
        doc.Save(outputPath);
        Console.WriteLine($"Document saved to {outputPath}");
    }
}
