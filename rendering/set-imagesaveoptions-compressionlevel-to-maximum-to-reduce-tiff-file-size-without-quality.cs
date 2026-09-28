using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Saving;

public class Program
{
    public static void Main()
    {
        // Create a simple document with some content.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);
        builder.Writeln("This is a sample document rendered to TIFF with maximum compression.");

        // Configure image save options for TIFF format.
        ImageSaveOptions saveOptions = new ImageSaveOptions(SaveFormat.Tiff);
        // Use the highest loss‑less compression available for TIFF.
        saveOptions.TiffCompression = TiffCompression.Lzw;

        // Define output file path.
        string outputPath = "output.tiff";

        // Save the document as a TIFF image using the configured options.
        doc.Save(outputPath, saveOptions);

        // Verify that the file was created.
        if (!File.Exists(outputPath))
        {
            throw new InvalidOperationException($"Failed to create the TIFF file at '{outputPath}'.");
        }

        // Report the file size to demonstrate compression effect.
        long fileSize = new FileInfo(outputPath).Length;
        Console.WriteLine($"TIFF file saved successfully. Size: {fileSize} bytes.");
    }
}
