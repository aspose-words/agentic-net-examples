using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Saving;

public class Program
{
    public static void Main()
    {
        // Define output folder and ensure it exists.
        string artifactsDir = Path.Combine(Directory.GetCurrentDirectory(), "Output");
        Directory.CreateDirectory(artifactsDir);

        // Create a simple DOCX document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);
        builder.Writeln("This is a sample document rendered to a color TIFF image.");
        builder.Writeln("Pixel format: 24 bits per pixel (RGB).");

        // Configure image save options for TIFF with 24bpp RGB pixel format.
        ImageSaveOptions saveOptions = new ImageSaveOptions(SaveFormat.Tiff)
        {
            PixelFormat = ImagePixelFormat.Format24BppRgb,
            // Optional: set resolution for better quality.
            Resolution = 300
        };

        // Save the document as a TIFF image.
        string tiffPath = Path.Combine(artifactsDir, "SampleDocument.tiff");
        doc.Save(tiffPath, saveOptions);

        // Verify that the TIFF file was created.
        if (!File.Exists(tiffPath))
            throw new InvalidOperationException("TIFF output file was not created.");

        // Indicate successful completion.
        Console.WriteLine("TIFF image saved successfully to: " + tiffPath);
    }
}
