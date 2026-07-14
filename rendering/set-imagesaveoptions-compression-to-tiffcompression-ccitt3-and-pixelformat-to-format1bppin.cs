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
        builder.Writeln("This is a sample document to be saved as a compressed TIFF image.");

        // Prepare the output folder.
        string artifactsDir = Path.Combine(Directory.GetCurrentDirectory(), "Artifacts");
        Directory.CreateDirectory(artifactsDir);

        // Define the output file path.
        string outputPath = Path.Combine(artifactsDir, "Compressed.tiff");

        // Configure image save options for smallest file size:
        // - Use CCITT3 compression (good for black‑and‑white images).
        // - Use 1‑bit per pixel indexed format.
        ImageSaveOptions options = new ImageSaveOptions(SaveFormat.Tiff)
        {
            TiffCompression = TiffCompression.Ccitt3,
            PixelFormat = ImagePixelFormat.Format1bppIndexed
        };

        // Save the document as a TIFF image using the configured options.
        doc.Save(outputPath, options);

        // Verify that the file was created.
        if (!File.Exists(outputPath))
            throw new InvalidOperationException("The TIFF file was not created.");

        // Optionally, output the file size (in bytes) to the console.
        long fileSize = new FileInfo(outputPath).Length;
        Console.WriteLine($"TIFF file saved successfully. Size: {fileSize} bytes.");
    }
}
