using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Saving;

public class Program
{
    public static void Main()
    {
        // Prepare output folder.
        string artifactsDir = Path.Combine(Directory.GetCurrentDirectory(), "Artifacts");
        Directory.CreateDirectory(artifactsDir);
        string outputPath = Path.Combine(artifactsDir, "output.tiff");

        // Create a simple document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);
        builder.Writeln("Sample text for TIFF conversion.");

        // Configure image save options for TIFF with 300 DPI.
        ImageSaveOptions options = new ImageSaveOptions(SaveFormat.Tiff);
        options.HorizontalResolution = 300f; // DpiX equivalent
        options.VerticalResolution = 300f;   // DpiY equivalent

        // Save the document as a TIFF image.
        doc.Save(outputPath, options);

        // Verify that the file was created.
        if (!File.Exists(outputPath))
            throw new Exception("TIFF file was not created.");

        // Indicate success (optional).
        Console.WriteLine("TIFF file saved successfully at: " + outputPath);
    }
}
