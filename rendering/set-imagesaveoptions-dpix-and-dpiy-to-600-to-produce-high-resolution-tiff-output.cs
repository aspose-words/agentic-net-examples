using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Saving;

public class Program
{
    public static void Main()
    {
        // Create a simple document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);
        builder.Writeln("This is a high‑resolution TIFF example.");

        // Configure image save options for TIFF with 600 DPI.
        ImageSaveOptions options = new ImageSaveOptions(SaveFormat.Tiff)
        {
            // Set both horizontal and vertical resolution to 600 DPI.
            HorizontalResolution = 600f,
            VerticalResolution = 600f
        };

        // Define output path and ensure the directory exists.
        string outputPath = Path.Combine(Directory.GetCurrentDirectory(), "HighResOutput.tiff");
        Directory.CreateDirectory(Path.GetDirectoryName(outputPath));

        // Save the document as a TIFF image.
        doc.Save(outputPath, options);

        // Verify that the file was created.
        if (!File.Exists(outputPath))
            throw new InvalidOperationException("Failed to create the TIFF file.");

        // Optionally, indicate success (no console input required).
        Console.WriteLine("TIFF file saved successfully at: " + outputPath);
    }
}
