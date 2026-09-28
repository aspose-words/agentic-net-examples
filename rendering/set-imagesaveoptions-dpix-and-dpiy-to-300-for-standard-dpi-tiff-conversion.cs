using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Saving;

public class Program
{
    public static void Main()
    {
        // Create a simple document with one paragraph.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);
        builder.Writeln("Hello, Aspose.Words TIFF rendering with 300 DPI.");

        // Path for the output TIFF file.
        string outputPath = "output.tiff";

        // Configure ImageSaveOptions for TIFF format with 300 DPI.
        ImageSaveOptions saveOptions = new ImageSaveOptions(SaveFormat.Tiff)
        {
            // The Resolution property sets both DpiX and DpiY.
            Resolution = 300
        };

        // Save the document as a TIFF image.
        doc.Save(outputPath, saveOptions);

        // Verify that the TIFF file was created.
        if (!File.Exists(outputPath))
        {
            throw new InvalidOperationException("Failed to create the TIFF file.");
        }

        // Output the location and size of the generated file.
        Console.WriteLine($"TIFF saved to: {Path.GetFullPath(outputPath)} (Size: {new FileInfo(outputPath).Length} bytes)");
    }
}
