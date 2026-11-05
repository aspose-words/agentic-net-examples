using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Saving;

public class Program
{
    public static void Main()
    {
        // Create a simple document in memory.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);
        builder.Writeln("Hello, high‑resolution TIFF!");

        // Configure image save options for TIFF with 600 DPI.
        // The ImageSaveOptions class uses the Resolution property to set both DpiX and DpiY.
        ImageSaveOptions saveOptions = new ImageSaveOptions(SaveFormat.Tiff)
        {
            Resolution = 600
        };

        // Define output file path.
        string outputPath = "high_res_output.tiff";

        // Save the document as a TIFF image.
        doc.Save(outputPath, saveOptions);

        // Verify that the file was created.
        if (!File.Exists(outputPath))
        {
            throw new InvalidOperationException($"Failed to create the TIFF file at '{outputPath}'.");
        }

        // Report success (no user interaction required).
        Console.WriteLine($"TIFF file saved successfully at '{Path.GetFullPath(outputPath)}' with 600 DPI.");
    }
}
