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

        // Insert a tiny placeholder PNG image (1x1 pixel) into the header.
        // The image is provided as a byte array to avoid dependencies on System.Drawing.
        byte[] pngData = Convert.FromBase64String(
            "iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAQAAAC1HAwCAAAAC0lEQVR42mP8/x8AAwMCAO+X2V8AAAAASUVORK5CYII=");
        Shape shape = builder.InsertImage(pngData);

        // Make the image float (not inline) and place it behind any text.
        shape.WrapType = WrapType.None;
        shape.BehindText = true;

        // Center the image on the page.
        shape.RelativeHorizontalPosition = RelativeHorizontalPosition.Page;
        shape.RelativeVerticalPosition = RelativeVerticalPosition.Page;
        shape.Left = (builder.PageSetup.PageWidth - shape.Width) / 2;
        shape.Top = (builder.PageSetup.PageHeight - shape.Height) / 2;

        // Return to the main document body and add some content.
        builder.MoveToSection(0);
        builder.Writeln("First page content.");
        builder.InsertBreak(BreakType.PageBreak);
        builder.Writeln("Second page content.");

        // Save the document to the local file system.
        doc.Save("FirstPageHeaderWithLogo.docx");
    }
}
