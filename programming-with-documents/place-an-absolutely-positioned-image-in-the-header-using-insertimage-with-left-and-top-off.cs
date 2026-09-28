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

        // Move to the primary header.
        builder.MoveToHeaderFooter(HeaderFooterType.HeaderPrimary);

        // A simple 1x1 PNG image (transparent pixel) encoded as Base64.
        const string base64Png =
            "iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAQAAAC1HAwCAAAAC0lEQVR42mP8/x8AAwMCAO+XG6cAAAAASUVORK5CYII=";

        // Convert the Base64 string to a byte array.
        byte[] imageBytes = Convert.FromBase64String(base64Png);

        using (MemoryStream imgStream = new MemoryStream(imageBytes))
        {
            // Insert the image with absolute positioning (left: 100 points, top: 50 points).
            // Width and height are set to the image's original size (approximately 1x1 points).
            builder.InsertImage(
                imgStream,
                RelativeHorizontalPosition.Page,
                100, // left offset in points
                RelativeVerticalPosition.Page,
                50,  // top offset in points
                100, // width in points
                50,  // height in points
                WrapType.None);
        }

        // Add some body text to the document.
        builder.MoveToDocumentEnd();
        builder.Writeln("This document contains an absolutely positioned image in the header.");

        // Save the document.
        string outputPath = "HeaderImage.docx";
        doc.Save(outputPath);

        // Verify that the file was created and can be reloaded.
        if (File.Exists(outputPath))
        {
            Document loadedDoc = new Document(outputPath);
            // Successful load confirms the file is valid.
        }
    }
}
