using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Saving;

public class Program
{
    public static void Main()
    {
        // Define output folder and ensure it exists.
        string artifactsDir = Path.Combine(Directory.GetCurrentDirectory(), "Artifacts");
        Directory.CreateDirectory(artifactsDir);

        // Create a simple Word document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);
        builder.Writeln("Hello World!");

        // Configure image save options for TIFF with custom DPI.
        ImageSaveOptions options = new ImageSaveOptions(SaveFormat.Tiff)
        {
            // Horizontal DPI (X axis) = 72.
            HorizontalResolution = 72f,
            // Vertical DPI (Y axis) = 144 (stretched vertically).
            VerticalResolution = 144f
        };

        // Save the document as a TIFF image.
        string outputPath = Path.Combine(artifactsDir, "VerticallyStretched.tiff");
        doc.Save(outputPath, options);

        // Verify that the file was created.
        if (!File.Exists(outputPath))
            throw new InvalidOperationException("The TIFF file was not created.");

        // Optionally, indicate success (no console interaction required).
    }
}
