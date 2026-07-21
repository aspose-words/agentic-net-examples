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
        builder.Writeln("Sample text for TIFF compression test.");

        // Configure image save options:
        // - Save as TIFF.
        // - Use CCITT3 compression (good for black‑and‑white images).
        // - Use 1‑bit per pixel indexed format to minimise file size.
        ImageSaveOptions options = new ImageSaveOptions(SaveFormat.Tiff)
        {
            TiffCompression = TiffCompression.Ccitt3,
            PixelFormat = ImagePixelFormat.Format1bppIndexed
        };

        // Define output path.
        string outputPath = Path.Combine(Directory.GetCurrentDirectory(), "output.tiff");

        // Save the document using the configured options.
        doc.Save(outputPath, options);

        // Verify that the file was created and is not empty.
        if (!File.Exists(outputPath))
            throw new InvalidOperationException("The TIFF file was not created.");

        long fileSize = new FileInfo(outputPath).Length;
        if (fileSize == 0)
            throw new InvalidOperationException("The TIFF file is empty.");
    }
}
