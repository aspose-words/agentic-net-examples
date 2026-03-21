using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Drawing;

class Program
{
    static void Main()
    {
        // Create a new empty document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // A 1x1 pixel PNG image (transparent) encoded in base64.
        const string base64Png = "iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAQAAAC1HAwCAAAAC0lEQVR42mP8/x8AAwMCAO+XK6cAAAAASUVORK5CYII=";
        byte[] pngBytes = Convert.FromBase64String(base64Png);

        // Insert the image shape into the document from the stream.
        using var imageStream = new MemoryStream(pngBytes);
        Shape shape = builder.InsertImage(imageStream);

        // Lock the shape's aspect ratio so that resizing preserves its proportions.
        shape.AspectRatioLocked = true;

        // Ensure the output directory exists.
        string outputDir = "ArtifactsDir";
        Directory.CreateDirectory(outputDir);

        // Save the resulting document.
        doc.Save(Path.Combine(outputDir, "Shape.AspectRatioLocked.docx"));
    }
}
