using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Saving;

class Program
{
    static void Main()
    {
        // Define output folder.
        string artifactsDir = Path.Combine(Environment.CurrentDirectory, "Artifacts");
        Directory.CreateDirectory(artifactsDir);

        // Create a new blank document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Add some sample content.
        builder.Writeln("Sample 1‑bit black‑white TIFF output.");
        // (Optional) Insert an image to demonstrate rendering.
        // builder.InsertImage("Path/To/Image.jpg");

        // Configure image save options for TIFF with 1‑bpp indexed pixel format.
        ImageSaveOptions saveOptions = new ImageSaveOptions(SaveFormat.Tiff)
        {
            // Set the pixel format to 1‑bit indexed (black‑white).
            PixelFormat = ImagePixelFormat.Format1bppIndexed,
            // Use CCITT Group 4 compression which is common for 1‑bpp TIFFs.
            TiffCompression = TiffCompression.Ccitt4
        };

        // Save the document as a TIFF image using the configured options.
        string outputPath = Path.Combine(artifactsDir, "Document_1bpp.tiff");
        doc.Save(outputPath, saveOptions);
    }
}
