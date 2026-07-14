using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Drawing;

public class Program
{
    public static void Main()
    {
        // Create a temporary image file to insert.
        string imagePath = Path.Combine(Directory.GetCurrentDirectory(), "sample.png");
        CreateSampleImage(imagePath);

        // Create a new document and a builder.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Insert the image as a shape.
        Shape shape = builder.InsertImage(imagePath);

        // Lock the shape's aspect ratio.
        shape.AspectRatioLocked = true;

        // Save the document.
        string docPath = Path.Combine(Directory.GetCurrentDirectory(), "AspectRatioLocked.docx");
        doc.Save(docPath);

        // Verify that the file was created.
        if (!File.Exists(docPath))
            throw new InvalidOperationException("The document was not saved successfully.");
    }

    // Creates a simple placeholder PNG image without using System.Drawing.
    private static void CreateSampleImage(string path)
    {
        // This is a 1x1 pixel transparent PNG encoded in base64.
        const string base64Png = "iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAQAAAC1HAwCAAAAC0lEQVR42mP8/x8AAwMCAO+XcZcAAAAASUVORK5CYII=";
        byte[] pngBytes = Convert.FromBase64String(base64Png);
        File.WriteAllBytes(path, pngBytes);
    }
}
