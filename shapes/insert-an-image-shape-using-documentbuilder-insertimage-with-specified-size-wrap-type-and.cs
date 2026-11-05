using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Drawing;

public class Program
{
    public static void Main()
    {
        // Create a temporary image file (a simple 1x1 PNG).
        string imagePath = "sample.png";
        CreateSampleImage(imagePath);

        // Create a new document and a builder.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Insert the image with explicit size (points).
        double width = 150;   // points
        double height = 150;  // points
        Shape imageShape = builder.InsertImage(imagePath, width, height);

        // Set wrapping behavior.
        imageShape.WrapType = WrapType.Square;

        // Position the shape relative to the page.
        imageShape.RelativeHorizontalPosition = RelativeHorizontalPosition.Page;
        imageShape.RelativeVerticalPosition = RelativeVerticalPosition.Page;
        imageShape.Left = 100; // points from the left edge
        imageShape.Top = 100;  // points from the top edge

        // Save the document.
        string docPath = "ImageShape.docx";
        doc.Save(docPath);

        // Validate that the document was saved.
        if (!File.Exists(docPath))
            throw new Exception("The output document was not created.");

        // Optional cleanup of the temporary image.
        // File.Delete(imagePath);
    }

    // Creates a minimal PNG file without using System.Drawing.
    private static void CreateSampleImage(string path)
    {
        // PNG data for a 1x1 pixel light blue image.
        byte[] pngData = new byte[]
        {
            0x89,0x50,0x4E,0x47,0x0D,0x0A,0x1A,0x0A,
            0x00,0x00,0x00,0x0D,0x49,0x48,0x44,0x52,
            0x00,0x00,0x00,0x01,0x00,0x00,0x00,0x01,
            0x08,0x02,0x00,0x00,0x00,0x90,0x77,0x53,
            0xDE,0x00,0x00,0x00,0x0A,0x49,0x44,0x41,
            0x54,0x08,0xD7,0x63,0xF8,0xCF,0xC0,0x00,
            0x00,0x04,0x00,0x01,0xE2,0x26,0x05,0x9B,
            0x00,0x00,0x00,0x00,0x49,0x45,0x4E,0x44,
            0xAE,0x42,0x60,0x82
        };
        File.WriteAllBytes(path, pngData);
    }
}
