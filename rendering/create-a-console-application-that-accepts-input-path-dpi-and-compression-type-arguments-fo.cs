using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Saving;

public class Program
{
    public static void Main(string[] args)
    {
        // Expect three arguments: input file path, DPI, and TIFF compression type.
        if (args.Length < 3)
        {
            Console.WriteLine("Usage: <app> <inputPath> <dpi> <compression>");
            return;
        }

        string inputPath = args[0];
        string dpiArg = args[1];
        string compressionArg = args[2];

        // Parse DPI.
        if (!float.TryParse(dpiArg, out float dpi) || dpi <= 0)
        {
            Console.WriteLine("Invalid DPI value.");
            return;
        }

        // Parse compression enum (case‑insensitive).
        if (!Enum.TryParse<TiffCompression>(compressionArg, true, out TiffCompression compression))
        {
            Console.WriteLine("Invalid compression type. Use None, Rle, Lzw, Ccitt3, or Ccitt4.");
            return;
        }

        // Ensure the source document exists; if not, create a simple one.
        if (!File.Exists(inputPath))
        {
            Document sampleDoc = new Document();
            DocumentBuilder builder = new DocumentBuilder(sampleDoc);
            builder.Writeln("This is a sample document created because the input file was missing.");
            sampleDoc.Save(inputPath);
        }

        // Load the document.
        Document doc = new Document(inputPath);

        // Configure TIFF save options.
        ImageSaveOptions options = new ImageSaveOptions(SaveFormat.Tiff)
        {
            Resolution = dpi,          // Sets both horizontal and vertical DPI.
            TiffCompression = compression
        };

        // Determine output file path (same name with .tiff extension).
        string outputPath = Path.ChangeExtension(inputPath, ".tiff");

        // Save the document as a TIFF image.
        doc.Save(outputPath, options);

        // Verify that the output file was created.
        if (!File.Exists(outputPath))
            throw new InvalidOperationException("TIFF conversion failed: output file not found.");

        // Optionally, report success.
        Console.WriteLine($"Document successfully converted to TIFF: {outputPath}");
    }
}
