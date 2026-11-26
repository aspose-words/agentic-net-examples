using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Saving;

public class Program
{
    public static void Main(string[] args)
    {
        // Default values used when command‑line arguments are not supplied.
        string inputPath = "sample.docx";
        int dpi = 300;
        TiffCompression compression = TiffCompression.Lzw;

        // If arguments are provided, they override the defaults.
        if (args.Length >= 3)
        {
            inputPath = args[0];

            if (!int.TryParse(args[1], out dpi) || dpi <= 0)
                throw new ArgumentException("DPI must be a positive integer.");

            if (!Enum.TryParse<TiffCompression>(args[2], true, out compression))
                throw new ArgumentException($"Invalid TIFF compression type: {args[2]}");
        }

        // Ensure the input document exists; create a simple sample if it does not.
        if (!File.Exists(inputPath))
        {
            var sampleDoc = new Document();
            var builder = new DocumentBuilder(sampleDoc);
            builder.Writeln("Sample document for TIFF conversion.");
            sampleDoc.Save(inputPath);
        }

        // Load the document.
        Document doc = new Document(inputPath);

        // Configure TIFF save options.
        ImageSaveOptions saveOptions = new ImageSaveOptions(SaveFormat.Tiff)
        {
            Resolution = dpi,
            TiffCompression = compression
        };

        // Determine output path (same folder, same name with .tiff extension).
        string outputPath = Path.ChangeExtension(inputPath, ".tiff");

        // Save the document as TIFF.
        doc.Save(outputPath, saveOptions);

        // Verify that the TIFF file was created.
        if (!File.Exists(outputPath))
            throw new InvalidOperationException("TIFF conversion failed; output file not found.");

        Console.WriteLine($"TIFF file created at: {outputPath}");
    }
}
