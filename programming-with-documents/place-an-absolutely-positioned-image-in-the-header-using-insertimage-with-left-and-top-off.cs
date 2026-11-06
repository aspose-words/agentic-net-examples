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

        // Move the builder to the primary header (creates it if missing).
        builder.MoveToHeaderFooter(HeaderFooterType.HeaderPrimary);

        // A tiny PNG image (1x1 pixel) encoded in base64 to avoid System.Drawing dependency.
        const string base64Png = "iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAQAAAC1HAwCAAAAC0lEQVR42mP8/x8AAwMCAO+XK6cAAAAASUVORK5CYII=";
        byte[] pngBytes = Convert.FromBase64String(base64Png);

        using (MemoryStream imageStream = new MemoryStream(pngBytes))
        {
            // Insert the image into the header.
            Shape shape = builder.InsertImage(imageStream);

            // Set the shape to be absolutely positioned with specific offsets.
            shape.WrapType = WrapType.None;                     // No text wrapping.
            shape.RelativeHorizontalPosition = RelativeHorizontalPosition.Page;
            shape.RelativeVerticalPosition = RelativeVerticalPosition.Page;

            // Offsets from the top‑left corner of the page (in points).
            shape.Left = ConvertUtil.MillimeterToPoint(10);     // 10 mm from the left.
            shape.Top = ConvertUtil.MillimeterToPoint(20);      // 20 mm from the top.
        }

        // Save the document to disk.
        const string outputPath = "HeaderImage.docx";
        doc.Save(outputPath);
    }
}
