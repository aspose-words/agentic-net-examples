using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Saving;

public class Program
{
    public static void Main(string[] args)
    {
        // If arguments are not supplied, use default sample values.
        // Expected arguments: <inputPath> <dpi> <compressionType>
        string inputPath;
        int dpi;
        TiffCompression compression;

        if (args.Length == 3)
        {
            inputPath = args[0];

            if (!int.TryParse(args[1], out dpi) || dpi <= 0)
                throw new ArgumentException("DPI must be a positive integer.");

            if (!Enum.TryParse<TiffCompression>(args[2], true, out compression))
                throw new ArgumentException("Invalid TIFF compression type.");
        }
        else
        {
            // Create a temporary sample document.
            string tempDir = Path.Combine(Path.GetTempPath(), "AsposeWordsTiffSample");
            Directory.CreateDirectory(tempDir);
            inputPath = Path.Combine(tempDir, "Sample.docx");
            dpi = 300;
            compression = TiffCompression.Lzw;
        }

        // Ensure the source document exists; create a simple one if it does not.
        if (!File.Exists(inputPath))
        {
            var doc = new Document();
            var builder = new DocumentBuilder(doc);
            builder.Writeln("Sample document for TIFF conversion.");
            doc.Save(inputPath);
        }

        // Load the document.
        Document document = new Document(inputPath);

        // Configure TIFF save options.
        ImageSaveOptions saveOptions = new ImageSaveOptions(SaveFormat.Tiff)
        {
            Resolution = dpi,
            TiffCompression = compression
        };

        // Determine output path (same folder, .tiff extension).
        string outputPath = Path.ChangeExtension(inputPath, ".tiff");

        // Save the document as TIFF.
        document.Save(outputPath, saveOptions);

        // Validate that the TIFF file was created.
        if (!File.Exists(outputPath))
            throw new InvalidOperationException("TIFF conversion failed; output file not found.");

        Console.WriteLine($"TIFF file saved to: {outputPath}");
    }
}
