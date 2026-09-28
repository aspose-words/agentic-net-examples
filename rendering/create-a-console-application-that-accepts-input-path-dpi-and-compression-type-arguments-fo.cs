using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Saving;

public class Program
{
    public static void Main(string[] args)
    {
        // Determine input path (use a default sample file if not provided)
        string inputPath = args.Length > 0 ? args[0] : "sample.docx";

        // Determine DPI (default to 300 if not provided or invalid)
        int dpi = 300;
        if (args.Length > 1 && int.TryParse(args[1], out int parsedDpi) && parsedDpi > 0)
            dpi = parsedDpi;

        // Determine compression type (default to Ccitt4 if not provided or invalid)
        string compressionStr = args.Length > 2 ? args[2] : "Ccitt4";

        // Ensure the input document exists; if not, create a simple one.
        if (!File.Exists(inputPath))
        {
            Document sampleDoc = new Document();
            DocumentBuilder builder = new DocumentBuilder(sampleDoc);
            builder.Writeln("Sample document for TIFF conversion.");
            sampleDoc.Save(inputPath);
        }

        // Load the document.
        Document doc = new Document(inputPath);

        // Prepare TIFF save options.
        ImageSaveOptions options = new ImageSaveOptions(SaveFormat.Tiff)
        {
            Resolution = dpi
        };

        // Parse compression type.
        if (Enum.TryParse<TiffCompression>(compressionStr, true, out TiffCompression compression))
        {
            options.TiffCompression = compression;
        }
        else
        {
            // Default to CCITT4 if parsing fails.
            options.TiffCompression = TiffCompression.Ccitt4;
        }

        // Determine output file path.
        string outputPath = Path.ChangeExtension(inputPath, ".tiff");

        // Save the document as TIFF.
        doc.Save(outputPath, options);

        // Validate that the output file was created.
        if (!File.Exists(outputPath))
            throw new InvalidOperationException($"Failed to create TIFF file at '{outputPath}'.");

        // Confirmation.
        Console.WriteLine($"TIFF conversion completed: '{outputPath}' (DPI={dpi}, Compression={options.TiffCompression})");
    }
}
