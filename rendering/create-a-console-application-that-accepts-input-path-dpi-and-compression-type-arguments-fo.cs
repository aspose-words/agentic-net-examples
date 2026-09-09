using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Saving;

public class Program
{
    public static void Main(string[] args)
    {
        // Default values – used when the required arguments are not supplied.
        const string defaultInput = "sample.docx";
        const int defaultDpi = 300;
        const TiffCompression defaultCompression = TiffCompression.Lzw;

        // Resolve input path.
        string inputPath = args.Length > 0 ? args[0] : defaultInput;

        // Resolve DPI.
        int dpi = defaultDpi;
        if (args.Length > 1 && int.TryParse(args[1], out int parsedDpi) && parsedDpi > 0)
            dpi = parsedDpi;

        // Resolve compression type.
        TiffCompression compression = defaultCompression;
        if (args.Length > 2 && Enum.TryParse<TiffCompression>(args[2], true, out TiffCompression parsedComp))
            compression = parsedComp;

        // Ensure the source document exists; create a simple one if it does not.
        if (!File.Exists(inputPath))
        {
            Document sampleDoc = new Document();
            DocumentBuilder builder = new DocumentBuilder(sampleDoc);
            builder.Writeln("Sample document generated because the input file was missing.");
            sampleDoc.Save(inputPath);
        }

        // Load the document.
        Document doc = new Document(inputPath);

        // Configure TIFF save options.
        ImageSaveOptions saveOptions = new ImageSaveOptions(SaveFormat.Tiff)
        {
            Resolution = dpi,          // Sets both horizontal and vertical DPI.
            TiffCompression = compression
        };

        // Determine output file name (same folder, same base name, .tiff extension).
        string outputPath = Path.ChangeExtension(inputPath, ".tiff");

        // Save as a (potentially multi‑page) TIFF image.
        doc.Save(outputPath, saveOptions);

        // Verify that the file was created.
        if (!File.Exists(outputPath))
            throw new InvalidOperationException($"Failed to create TIFF file at '{outputPath}'.");

        Console.WriteLine($"TIFF conversion completed: {outputPath}");
    }
}
