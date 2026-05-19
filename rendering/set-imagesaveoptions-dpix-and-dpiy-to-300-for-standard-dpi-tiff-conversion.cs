using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Saving;

public class Program
{
    public static void Main()
    {
        // Define output file path.
        string outputPath = Path.Combine(Directory.GetCurrentDirectory(), "StandardDPI.tiff");

        // Create a new document and add some content.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);
        builder.Writeln("This is a sample document rendered to TIFF with 300 DPI.");

        // Configure image save options for TIFF format.
        ImageSaveOptions options = new ImageSaveOptions(SaveFormat.Tiff)
        {
            // Set horizontal and vertical resolution to 300 DPI.
            HorizontalResolution = 300f,
            VerticalResolution = 300f
        };

        // Save the document as a TIFF image.
        doc.Save(outputPath, options);

        // Verify that the file was created.
        if (!File.Exists(outputPath))
            throw new InvalidOperationException($"Failed to create the TIFF file at '{outputPath}'.");

        // Optionally, output the file size to confirm creation.
        Console.WriteLine($"TIFF file saved successfully. Size: {new FileInfo(outputPath).Length} bytes.");
    }
}
