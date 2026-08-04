using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Saving;

public class Program
{
    public static void Main(string[] args)
    {
        // Expected arguments: <inputPath> <dpi> <compression>
        // If arguments are missing or invalid, defaults are used.
        string inputPath = args.Length > 0 ? args[0] : string.Empty;
        int dpi = args.Length > 1 && int.TryParse(args[1], out int parsedDpi) ? parsedDpi : 300;
        string compressionArg = args.Length > 2 ? args[2] : "Lzw";

        // Resolve compression type.
        if (!Enum.TryParse<TiffCompression>(compressionArg, true, out TiffCompression compression))
        {
            compression = TiffCompression.Lzw; // fallback to default
        }

        // Ensure we have a source document.
        Document doc;
        if (!string.IsNullOrEmpty(inputPath) && File.Exists(inputPath))
        {
            doc = new Document(inputPath);
        }
        else
        {
            // Create a simple sample document.
            doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);
            builder.Writeln("Sample document for TIFF conversion.");
            // Save the sample document locally so it can be reloaded if needed.
            string samplePath = Path.Combine(Path.GetTempPath(), "SampleDocument.docx");
            doc.Save(samplePath);
            doc = new Document(samplePath);
        }

        // Prepare TIFF save options.
        ImageSaveOptions tiffOptions = new ImageSaveOptions(SaveFormat.Tiff)
        {
            Resolution = dpi,
            TiffCompression = compression
        };

        // Determine output file path.
        string outputDirectory = Path.GetDirectoryName(inputPath);
        if (string.IsNullOrEmpty(outputDirectory) || !Directory.Exists(outputDirectory))
            outputDirectory = Directory.GetCurrentDirectory();

        string outputFileName = $"Converted_{dpi}dpi_{compression}.tiff";
        string outputPath = Path.Combine(outputDirectory, outputFileName);

        // Save the document as a TIFF image.
        doc.Save(outputPath, tiffOptions);

        // Validate that the file was created.
        if (!File.Exists(outputPath))
            throw new InvalidOperationException($"Failed to create TIFF file at '{outputPath}'.");

        // Optionally, inform the user (no interactive wait).
        Console.WriteLine($"TIFF file saved to: {outputPath}");
    }
}
