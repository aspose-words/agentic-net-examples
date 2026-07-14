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
        string outputPath = Path.Combine(artifactsDir, "HighRes.tiff");

        // Create a simple document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);
        builder.Writeln("High resolution TIFF example.");

        // Configure image save options for 600 DPI.
        ImageSaveOptions options = new ImageSaveOptions(SaveFormat.Tiff)
        {
            HorizontalResolution = 600f,
            VerticalResolution = 600f
        };

        // Save the document as a TIFF image.
        doc.Save(outputPath, options);

        // Verify that the file was created.
        if (!File.Exists(outputPath))
            throw new InvalidOperationException("The TIFF file was not created.");

        // Optionally, indicate success (no console input required).
        Console.WriteLine("TIFF saved successfully at: " + outputPath);
    }
}
