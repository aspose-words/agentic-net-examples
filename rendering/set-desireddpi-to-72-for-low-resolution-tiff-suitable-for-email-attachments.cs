using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Saving;

public class Program
{
    public static void Main()
    {
        // Create a simple document with some text.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);
        builder.Writeln("This is a low‑resolution TIFF image generated for email attachments.");

        // Configure image save options for TIFF.
        ImageSaveOptions options = new ImageSaveOptions(SaveFormat.Tiff);

        // Set the resolution (DPI) to 72. This property sets both horizontal and vertical DPI.
        options.Resolution = 72f;

        // Optionally, set the explicit horizontal and vertical resolution properties as well.
        options.HorizontalResolution = 72f;
        options.VerticalResolution = 72f;

        // Define the output file path.
        string outputPath = "LowResolution.tiff";

        // Save the document as a TIFF image using the configured options.
        doc.Save(outputPath, options);

        // Verify that the file was created.
        if (!File.Exists(outputPath))
            throw new InvalidOperationException("Failed to create the TIFF file.");

        // Inform the user where the file was saved.
        Console.WriteLine($"TIFF saved successfully to: {Path.GetFullPath(outputPath)}");
    }
}
