using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Saving;

public class Program
{
    public static void Main(string[] args)
    {
        // Resolve input file path.
        string inputPath;
        if (args.Length >= 1 && File.Exists(args[0]))
        {
            inputPath = args[0];
        }
        else
        {
            // No valid input file – create a temporary sample document.
            inputPath = Path.Combine(Directory.GetCurrentDirectory(), "SampleDocument.docx");
            Document sampleDoc = new Document();
            DocumentBuilder builder = new DocumentBuilder(sampleDoc);
            builder.Writeln("This is a sample document generated for TIFF conversion.");

            // Insert a placeholder image only if one exists.
            string placeholder = GetPlaceholderImagePath();
            if (!string.IsNullOrEmpty(placeholder))
                builder.InsertImage(placeholder);

            sampleDoc.Save(inputPath);
        }

        // Resolve DPI (default 300).
        int dpi = 300;
        if (args.Length >= 2 && int.TryParse(args[1], out int parsedDpi) && parsedDpi > 0)
            dpi = parsedDpi;

        // Resolve compression type (default Lzw).
        TiffCompression compression = TiffCompression.Lzw;
        if (args.Length >= 3 && Enum.TryParse<TiffCompression>(args[2], true, out TiffCompression parsedComp))
            compression = parsedComp;

        // Load the source document.
        Document doc = new Document(inputPath);

        // Configure TIFF save options.
        ImageSaveOptions saveOptions = new ImageSaveOptions(SaveFormat.Tiff)
        {
            Resolution = dpi,
            TiffCompression = compression
        };

        // Determine output file path.
        string outputPath = Path.ChangeExtension(inputPath, ".tiff");

        // Save the document as a TIFF image.
        doc.Save(outputPath, saveOptions);

        // Verify that the output file was created.
        if (!File.Exists(outputPath))
            throw new InvalidOperationException($"Failed to create TIFF file: {outputPath}");
    }

    // Attempts to locate a sample image in the current directory.
    private static string GetPlaceholderImagePath()
    {
        string[] possibleNames = { "Placeholder.png", "Placeholder.jpg", "Placeholder.bmp" };
        foreach (var name in possibleNames)
        {
            string candidate = Path.Combine(Directory.GetCurrentDirectory(), name);
            if (File.Exists(candidate))
                return candidate;
        }
        return string.Empty;
    }
}
