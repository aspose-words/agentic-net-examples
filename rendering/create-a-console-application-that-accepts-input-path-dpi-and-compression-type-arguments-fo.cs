using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Saving;

public class Program
{
    public static void Main(string[] args)
    {
        // Input document path (first argument) or default sample path.
        string inputPath = args.Length > 0 ? args[0] : "Sample.docx";

        // DPI (second argument) or default 300.
        int dpi = 300;
        if (args.Length > 1 && int.TryParse(args[1], out int parsedDpi) && parsedDpi > 0)
            dpi = parsedDpi;

        // Compression type (third argument) or default Lzw.
        string compressionArg = args.Length > 2 ? args[2] : "Lzw";
        if (!Enum.TryParse<TiffCompression>(compressionArg, true, out TiffCompression compression))
            compression = TiffCompression.Lzw; // fallback to Lzw if parsing fails.

        // Ensure the source document exists; create a simple one if it does not.
        if (!File.Exists(inputPath))
        {
            Document sampleDoc = new Document();
            DocumentBuilder builder = new DocumentBuilder(sampleDoc);
            builder.Writeln("This is a sample document generated for TIFF conversion.");
            builder.InsertImage(GetPlaceholderImagePath());
            sampleDoc.Save(inputPath);
        }

        // Load the document.
        Document doc = new Document(inputPath);

        // Configure TIFF rendering options.
        ImageSaveOptions options = new ImageSaveOptions(SaveFormat.Tiff)
        {
            Resolution = dpi,
            TiffCompression = compression
        };

        // Determine output file path (same name with .tiff extension).
        string outputPath = Path.ChangeExtension(inputPath, "tiff");

        // Save the document as a TIFF image.
        doc.Save(outputPath, options);

        // Verify that the output file was created.
        if (!File.Exists(outputPath))
            throw new InvalidOperationException($"Failed to create TIFF file at '{outputPath}'.");

        // Optional: indicate success (no interactive input required).
        Console.WriteLine($"Document converted to TIFF successfully: {outputPath}");
    }

    // Helper to provide a placeholder image for the sample document.
    private static string GetPlaceholderImagePath()
    {
        // Create a temporary PNG image using Aspose.Words (no external libraries).
        string tempImagePath = Path.Combine(Path.GetTempPath(), "placeholder.png");
        if (!File.Exists(tempImagePath))
        {
            // Generate a simple 1x1 pixel PNG using Aspose.Words' built‑in capabilities.
            // Since Aspose.Words does not create images directly, we embed a small base64 PNG.
            const string base64Png = "iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAQAAAC1HAwCAAAAC0lEQVR42mP8Xw8AApEB9V6XK6cAAAAASUVORK5CYII=";
            byte[] pngBytes = Convert.FromBase64String(base64Png);
            File.WriteAllBytes(tempImagePath, pngBytes);
        }
        return tempImagePath;
    }
}
