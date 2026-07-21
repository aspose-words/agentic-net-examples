using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Saving;

public class Program
{
    public static void Main()
    {
        // Define output file path.
        string outputPath = "BlackWhite.tiff";

        // Create a new blank document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Add some sample content.
        builder.Writeln("This is a sample document rendered as a black‑and‑white TIFF.");

        // Configure image save options for TIFF with 1‑bpp indexed pixel format.
        ImageSaveOptions saveOptions = new ImageSaveOptions(SaveFormat.Tiff)
        {
            PixelFormat = ImagePixelFormat.Format1bppIndexed,
            // Optional: use CCITT compression suitable for 1‑bpp images.
            TiffCompression = TiffCompression.Ccitt4
        };

        // Save the document as a TIFF image.
        doc.Save(outputPath, saveOptions);

        // Verify that the file was created.
        if (!File.Exists(outputPath))
            throw new InvalidOperationException("The TIFF file was not created.");

        // Optionally, you could output a confirmation (not required for the task).
        // Console.WriteLine($"TIFF saved to {Path.GetFullPath(outputPath)}");
    }
}
