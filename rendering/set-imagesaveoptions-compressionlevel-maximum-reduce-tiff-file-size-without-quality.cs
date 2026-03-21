using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Saving;

class Program
{
    static void Main()
    {
        // Create a new document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Insert a tiny 1x1 PNG image from an in‑memory stream.
        // This avoids the need for an external image file.
        const string base64Png = "iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAQAAAC1HAwCAAAAC0lEQVR42mP8/x8AAwMCAO+XK6cAAAAASUVORK5CYII=";
        byte[] imageBytes = Convert.FromBase64String(base64Png);
        using var imageStream = new MemoryStream(imageBytes);
        builder.InsertImage(imageStream);

        // Configure TIFF save options with maximum loss‑less compression.
        ImageSaveOptions tiffOptions = new ImageSaveOptions(SaveFormat.Tiff)
        {
            TiffCompression = TiffCompression.Lzw
        };

        // Save the document as a TIFF image.
        doc.Save("output.tiff", tiffOptions);
    }
}
