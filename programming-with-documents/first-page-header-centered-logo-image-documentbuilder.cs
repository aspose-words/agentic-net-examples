using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Drawing;

class Program
{
    static void Main()
    {
        // Create a new blank document.
        Document doc = new Document();

        // Attach a DocumentBuilder to the document.
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Enable a different header/footer for the first page.
        builder.PageSetup.DifferentFirstPageHeaderFooter = true;

        // Move the cursor to the first-page header.
        builder.MoveToHeaderFooter(HeaderFooterType.HeaderFirst);

        // Create a tiny placeholder PNG image in memory.
        const string base64Png = "iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAQAAAC1HAwCAAAAC0lEQVR42mP8/x8AAwMCAO+XbZcAAAAASUVORK5CYII=";
        byte[] pngBytes = Convert.FromBase64String(base64Png);
        using var imageStream = new MemoryStream(pngBytes);

        // Insert the image from the stream.
        Shape logo = builder.InsertImage(imageStream);

        // Configure the image as a floating shape behind the text.
        logo.WrapType = WrapType.None;
        logo.BehindText = true;

        // Position the image relative to the page.
        logo.RelativeHorizontalPosition = RelativeHorizontalPosition.Page;
        logo.RelativeVerticalPosition = RelativeVerticalPosition.Page;

        // Center the image horizontally and vertically on the page.
        logo.HorizontalAlignment = HorizontalAlignment.Center;
        logo.VerticalAlignment = VerticalAlignment.Center;

        // Return the cursor to the main document body.
        builder.MoveToDocumentEnd();

        // Add sample content to demonstrate the header.
        builder.Writeln("This is the first page.");
        builder.InsertBreak(BreakType.PageBreak);
        builder.Writeln("This is the second page.");

        // Save the document.
        doc.Save("FirstPageHeader.docx");
    }
}
