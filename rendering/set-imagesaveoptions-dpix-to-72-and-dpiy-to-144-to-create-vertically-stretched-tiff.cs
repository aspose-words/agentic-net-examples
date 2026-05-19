using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Saving;

public class Program
{
    public static void Main()
    {
        // Define output directory and ensure it exists.
        string artifactsDir = Path.Combine(Directory.GetCurrentDirectory(), "Artifacts");
        Directory.CreateDirectory(artifactsDir);
        string outputPath = Path.Combine(artifactsDir, "VerticallyStretched.tiff");

        // Create a simple document with two pages.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);
        builder.Writeln("This is a sample document.");
        builder.InsertBreak(BreakType.PageBreak);
        builder.Writeln("Second page.");

        // Configure image save options for TIFF with different horizontal and vertical DPI.
        ImageSaveOptions options = new ImageSaveOptions(SaveFormat.Tiff);
        options.HorizontalResolution = 72f; // DpiX
        options.VerticalResolution = 144f;   // DpiY

        // Save the document as a TIFF image using the specified options.
        doc.Save(outputPath, options);

        // Verify that the file was created.
        if (!File.Exists(outputPath))
            throw new InvalidOperationException("The TIFF file was not created.");

        // Optionally, you could output a confirmation (no interactive prompts required).
        Console.WriteLine("TIFF image saved successfully to: " + outputPath);
    }
}
