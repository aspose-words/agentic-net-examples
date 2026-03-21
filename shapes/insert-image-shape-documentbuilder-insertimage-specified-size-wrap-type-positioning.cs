using Aspose.Words;
using Aspose.Words.Drawing;
using System;
using System.IO;

class Program
{
    static void Main()
    {
        // Create a new empty document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Create a temporary PNG image (1x1 pixel, transparent).
        string tempImagePath = Path.Combine(Path.GetTempPath(), $"temp_{Guid.NewGuid()}.png");
        byte[] pngBytes = Convert.FromBase64String(
            "iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAQAAAC1HAwCAAAAC0lEQVR42mP8/x8AAwMCAO+XK6cAAAAASUVORK5CYII=");
        File.WriteAllBytes(tempImagePath, pngBytes);

        // Insert a floating image with custom size, position and wrap type.
        // - Position is relative to the page margins.
        // - Left offset: 100 points, Top offset: 50 points.
        // - Width: 200 points, Height: 150 points.
        // - Text will wrap around the image in a square fashion.
        Shape imageShape = builder.InsertImage(
            tempImagePath,
            RelativeHorizontalPosition.Margin, 100,
            RelativeVerticalPosition.Margin, 50,
            200, 150,
            WrapType.Square);

        // Example of an additional property: keep the image in front of the text.
        imageShape.BehindText = false;

        // Ensure output directory exists and save the document.
        string outputPath = Path.Combine(Directory.GetCurrentDirectory(), "ImageInserted.docx");
        doc.Save(outputPath);

        // Clean up temporary image file.
        File.Delete(tempImagePath);
    }
}
