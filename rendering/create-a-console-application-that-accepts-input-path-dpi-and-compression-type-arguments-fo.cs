using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Saving;

public class Program
{
    public static void Main(string[] args)
    {
        // Default values for demonstration when arguments are not supplied.
        string inputPath = args.Length > 0 ? args[0] : "SampleDocument.docx";
        int dpi = args.Length > 1 && int.TryParse(args[1], out int parsedDpi) && parsedDpi > 0
            ? parsedDpi
            : 300; // Default DPI

        // Parse compression type; default to Lzw if parsing fails or argument missing.
        TiffCompression compression = TiffCompression.Lzw;
        if (args.Length > 2 && Enum.TryParse<TiffCompression>(args[2], true, out TiffCompression parsedComp))
            compression = parsedComp;

        // Ensure the source document exists; create a simple one if it does not.
        if (!File.Exists(inputPath))
        {
            Document sampleDoc = new Document();
            DocumentBuilder builder = new DocumentBuilder(sampleDoc);
            builder.Writeln("Sample document for TIFF conversion.");
            sampleDoc.Save(inputPath);
        }

        // Load the document.
        Document doc = new Document(inputPath);

        // Configure TIFF save options.
        ImageSaveOptions options = new ImageSaveOptions(SaveFormat.Tiff)
        {
            Resolution = dpi,
            TiffCompression = compression
        };

        // Determine output file path (same directory, same name with .tiff extension).
        string outputPath = Path.ChangeExtension(inputPath, ".tiff");

        // Save the document as a TIFF image.
        doc.Save(outputPath, options);

        // Verify that the output file was created.
        if (!File.Exists(outputPath))
            throw new InvalidOperationException("TIFF conversion failed; output file not found.");

        Console.WriteLine($"TIFF image saved to: {outputPath}");
    }
}
