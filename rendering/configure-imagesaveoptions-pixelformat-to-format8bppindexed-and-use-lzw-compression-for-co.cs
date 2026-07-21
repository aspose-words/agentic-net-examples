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
        builder.Writeln("This is a sample document rendered to a color TIFF image.");
        builder.InsertBreak(BreakType.PageBreak);
        builder.Writeln("Second page of the document.");

        // Configure image save options for a color TIFF with LZW compression.
        ImageSaveOptions options = new ImageSaveOptions(SaveFormat.Tiff)
        {
            // Use a 24‑bit RGB pixel format (closest supported format for color output).
            PixelFormat = ImagePixelFormat.Format24BppRgb,
            // Apply LZW compression.
            TiffCompression = TiffCompression.Lzw
        };

        // Define output file path.
        string outputPath = Path.Combine(Directory.GetCurrentDirectory(), "Sample.tiff");

        // Save the document as a TIFF image using the configured options.
        doc.Save(outputPath, options);

        // Verify that the file was created.
        if (!File.Exists(outputPath))
            throw new InvalidOperationException("The TIFF file was not created.");

        // Report success.
        Console.WriteLine("TIFF image saved successfully to: " + outputPath);
    }
}
