using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Settings;
using Aspose.Words.Drawing;

public class Program
{
    public static void Main()
    {
        // Define paths for the temporary image and the output document.
        string imagePath = Path.Combine(Directory.GetCurrentDirectory(), "SampleImage.png");
        string outputPath = Path.Combine(Directory.GetCurrentDirectory(), "WatermarkedDocument.docx");

        // Create a minimal 1x1 PNG image from a Base64 string and write it to disk.
        // This avoids any System.Drawing or SkiaSharp usage.
        const string pngBase64 = "iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAQAAAC1HAwCAAAAC0lEQVR42mP8/5+BFwAE/wJ/6VZcAAAAAElFTkSuQmCC";
        byte[] pngBytes = Convert.FromBase64String(pngBase64);
        File.WriteAllBytes(imagePath, pngBytes);

        // Create a new blank document and add a simple paragraph.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);
        builder.Writeln("This document contains an image watermark.");

        // Optimize the document for Word 2010 compatibility.
        doc.CompatibilityOptions.OptimizeFor(MsWordVersion.Word2010);

        // Insert the image watermark using the file path.
        // Use the overload that accepts a path and ImageWatermarkOptions (options can be default).
        doc.Watermark.SetImage(imagePath, new ImageWatermarkOptions());

        // Save the resulting document.
        doc.Save(outputPath);
    }
}
