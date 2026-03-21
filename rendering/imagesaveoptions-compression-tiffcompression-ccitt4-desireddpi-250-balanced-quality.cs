using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Saving;

class Program
{
    static void Main()
    {
        // Create a new document and add some content.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);
        builder.Writeln("Sample text for black‑and‑white TIFF.");

        // Insert a simple 1x1 PNG image from a base64 string.
        const string pngBase64 = "iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAQAAAC1HAwCAAAAC0lEQVR42mP8/x8AAwMCAO+XK6cAAAAASUVORK5CYII=";
        byte[] pngBytes = Convert.FromBase64String(pngBase64);
        using (var ms = new MemoryStream(pngBytes))
        {
            builder.InsertImage(ms);
        }

        // Configure save options for a balanced‑quality B&W TIFF.
        ImageSaveOptions options = new ImageSaveOptions(SaveFormat.Tiff)
        {
            // Use CCITT4 compression, which is efficient for 1‑bpp images.
            TiffCompression = TiffCompression.Ccitt4,
            // Set both horizontal and vertical DPI to 250.
            Resolution = 250,
            // Render the pages as black‑and‑white.
            ImageColorMode = ImageColorMode.BlackAndWhite
        };

        // Save the document as a TIFF file using the configured options.
        doc.Save("output.tiff", options);
    }
}
