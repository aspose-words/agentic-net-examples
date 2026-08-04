using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Saving;

public class Program
{
    public static void Main()
    {
        // Prepare output folder.
        string outputDir = Path.Combine(Directory.GetCurrentDirectory(), "Output");
        Directory.CreateDirectory(outputDir);
        string outputPath = Path.Combine(outputDir, "VerticallyStretched.tiff");

        // Create a simple document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);
        builder.Writeln("Sample text for a vertically stretched TIFF image.");

        // Configure image save options: set horizontal DPI to 72 and vertical DPI to 144.
        ImageSaveOptions options = new ImageSaveOptions(SaveFormat.Tiff)
        {
            HorizontalResolution = 72f,
            VerticalResolution = 144f
        };

        // Save the document as a TIFF image using the specified DPI settings.
        doc.Save(outputPath, options);

        // Verify that the file was created.
        if (!File.Exists(outputPath))
            throw new InvalidOperationException("The TIFF file was not created.");

        // Optionally, indicate success (no interactive output required).
        Console.WriteLine("TIFF image saved successfully to: " + outputPath);
    }
}
