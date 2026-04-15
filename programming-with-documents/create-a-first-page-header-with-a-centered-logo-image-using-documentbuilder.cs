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

        // Enable a different header for the first page.
        builder.PageSetup.DifferentFirstPageHeaderFooter = true;

        // Move the cursor to the first‑page header.
        builder.MoveToHeaderFooter(HeaderFooterType.HeaderFirst);

        // A tiny 1x1 PNG image (transparent) encoded in Base64.
        // This avoids the need for System.Drawing or external image files.
        string base64Png = "iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAQAAAC1HAwCAAAAC0lEQVR42mP8/x8AAwMCAO+XK6cAAAAASUVORK5CYII=";
        byte[] pngBytes = Convert.FromBase64String(base64Png);

        // Insert the image into the header.
        Shape shape = builder.InsertImage(pngBytes);

        // Configure the shape to be centered on the page.
        shape.WrapType = WrapType.None;
        shape.BehindText = true;
        shape.RelativeHorizontalPosition = RelativeHorizontalPosition.Page;
        shape.RelativeVerticalPosition = RelativeVerticalPosition.Page;
        shape.HorizontalAlignment = HorizontalAlignment.Center;
        shape.VerticalAlignment = VerticalAlignment.Center;

        // Return to the main document body and add some sample text.
        builder.MoveToSection(0);
        builder.Writeln("Document with a first‑page header containing a centered logo.");

        // Save the document to the local file system.
        const string outputPath = "FirstPageHeaderWithLogo.docx";
        doc.Save(outputPath);
    }
}
