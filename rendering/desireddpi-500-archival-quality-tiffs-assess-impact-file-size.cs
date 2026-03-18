using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Saving;

class Program
{
    static void Main()
    {
        // Create a simple document with some text and an image.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);
        builder.Writeln("Archival‑quality TIFF test.");
        builder.Writeln("This document will be rendered at 500 DPI.");
        // Insert an image to increase the amount of raster data.
        // Replace with a valid image path on your system.
        builder.InsertImage("ImageDir/Logo.jpg");

        // Configure TIFF save options.
        ImageSaveOptions tiffOptions = new ImageSaveOptions(SaveFormat.Tiff)
        {
            // Desired resolution for archival quality.
            Resolution = 500f,

            // Use high‑quality (slow) rendering algorithms.
            UseHighQualityRendering = true,

            // Lossless LZW compression keeps image data intact.
            TiffCompression = TiffCompression.Lzw,

            // Render each page as a separate frame in a multi‑page TIFF.
            PageLayout = MultiPageLayout.TiffFrames()
        };

        // Save the document as a TIFF file.
        string tiffPath = "ArtifactsDir/Archival_500dpi.tiff";
        doc.Save(tiffPath, tiffOptions);

        // Assess the resulting file size.
        long fileSize = new FileInfo(tiffPath).Length;
        Console.WriteLine($"Saved TIFF size: {fileSize} bytes");
    }
}
