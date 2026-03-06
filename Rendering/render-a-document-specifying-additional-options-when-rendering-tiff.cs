using System;
using System.Drawing;
using Aspose.Words;
using Aspose.Words.Saving;

class RenderTiffWithOptions
{
    static void Main()
    {
        // Folder where the output file will be saved.
        string artifactsDir = "Output/";
        System.IO.Directory.CreateDirectory(artifactsDir);

        // Create a new blank document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Add some content to the document.
        builder.Font.Name = "Times New Roman";
        builder.Font.Size = 24;
        builder.Writeln("Sample text for TIFF rendering with additional options.");
        builder.InsertImage("ImageDir/Logo.jpg"); // Replace with actual image path.

        // Create ImageSaveOptions for TIFF format.
        ImageSaveOptions tiffOptions = new ImageSaveOptions(SaveFormat.Tiff);

        // Set TIFF-specific compression.
        tiffOptions.TiffCompression = TiffCompression.Ccitt4;

        // Configure binarization method and threshold for Floyd‑Steinberg dithering.
        tiffOptions.TiffBinarizationMethod = ImageBinarizationMethod.FloydSteinbergDithering;
        tiffOptions.ThresholdForFloydSteinbergDithering = 200; // Range 0‑255.

        // Set resolution (both horizontal and vertical) to 300 DPI.
        tiffOptions.Resolution = 300;

        // Optional: set background (paper) color.
        tiffOptions.PaperColor = Color.White;

        // Enable high‑quality rendering.
        tiffOptions.UseAntiAliasing = true;
        tiffOptions.UseHighQualityRendering = true;

        // Save the document as a TIFF image using the configured options.
        doc.Save(artifactsDir + "DocumentWithTiffOptions.tiff", tiffOptions);
    }
}
