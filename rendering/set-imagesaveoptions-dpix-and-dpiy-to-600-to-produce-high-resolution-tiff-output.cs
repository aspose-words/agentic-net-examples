using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Saving;

public class Program
{
    public static void Main()
    {
        // Create a new blank document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Add some sample content.
        builder.Writeln("Hello World! This is a sample document for high‑resolution TIFF rendering.");

        // Configure image save options for TIFF with 600 DPI.
        // In Aspose.Words the DPI is set via the Resolution property (applies to both X and Y).
        ImageSaveOptions saveOptions = new ImageSaveOptions(SaveFormat.Tiff)
        {
            Resolution = 600
        };

        // Define output file path.
        string outputPath = "HighResOutput.tiff";

        // Save the document as a high‑resolution TIFF image.
        doc.Save(outputPath, saveOptions);

        // Verify that the file was created.
        if (!File.Exists(outputPath))
        {
            throw new Exception($"Failed to create the TIFF file at '{outputPath}'.");
        }

        // Indicate success.
        Console.WriteLine($"TIFF file saved successfully at '{Path.GetFullPath(outputPath)}'.");
    }
}
