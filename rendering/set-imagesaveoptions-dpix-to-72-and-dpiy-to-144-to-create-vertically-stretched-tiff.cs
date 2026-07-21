using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Saving;

public class Program
{
    public static void Main()
    {
        // Create a simple document with a few lines of text.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);
        builder.Writeln("First line of text.");
        builder.Writeln("Second line of text.");
        builder.Writeln("Third line of text.");

        // Configure image save options for TIFF with different horizontal and vertical DPI.
        ImageSaveOptions saveOptions = new ImageSaveOptions(SaveFormat.Tiff);
        saveOptions.HorizontalResolution = 72f; // DpiX
        saveOptions.VerticalResolution = 144f;  // DpiY

        // Define the output file path.
        string outputDir = Path.Combine(Directory.GetCurrentDirectory(), "Output");
        Directory.CreateDirectory(outputDir);
        string outputPath = Path.Combine(outputDir, "VerticallyStretched.tiff");

        // Save the document as a TIFF image using the specified DPI settings.
        doc.Save(outputPath, saveOptions);

        // Verify that the file was created.
        if (!File.Exists(outputPath))
            throw new InvalidOperationException("The TIFF file was not created.");

        // Optionally, inform that the process completed successfully.
        Console.WriteLine("TIFF image saved to: " + outputPath);
    }
}
