using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Saving;

public class Program
{
    public static void Main()
    {
        // Define output directory.
        string outputDir = Path.Combine(Directory.GetCurrentDirectory(), "Output");
        Directory.CreateDirectory(outputDir);

        // Create a new blank document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Write a line with normal characters.
        builder.Writeln("The quick brown fox jumps over the lazy dog.");

        // Write a line that contains a discretionary ligature (ﬁ).
        // Unicode character U+FB01 is the ligature for "fi".
        builder.Writeln("Discretionary ligature test: ﬁ");

        // Configure image save options for TIFF.
        ImageSaveOptions saveOptions = new ImageSaveOptions(SaveFormat.Tiff)
        {
            // Use high quality rendering to preserve typographic details.
            UseAntiAliasing = true,
            UseHighQualityRendering = true
            // Do not set PageSet – leaving it null renders all pages into a single multi‑page TIFF.
        };

        // Save the document as a TIFF image.
        string tiffPath = Path.Combine(outputDir, "DocumentWithLigature.tiff");
        doc.Save(tiffPath, saveOptions);

        // Verify that the TIFF file was created and has content.
        if (!File.Exists(tiffPath))
            throw new FileNotFoundException("TIFF file was not created.", tiffPath);

        byte[] tiffBytes = File.ReadAllBytes(tiffPath);
        if (tiffBytes.Length == 0)
            throw new InvalidDataException("TIFF file is empty.");

        // Simple confirmation output.
        Console.WriteLine($"TIFF file generated successfully at: {tiffPath}");
        Console.WriteLine($"File size: {tiffBytes.Length} bytes");
    }
}
