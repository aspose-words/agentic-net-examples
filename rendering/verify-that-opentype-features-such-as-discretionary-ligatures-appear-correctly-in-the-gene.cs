using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Saving;

public class Program
{
    public static void Main()
    {
        // Create a new blank document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Choose a font that supports ligatures.
        builder.Font.Name = "Arial Unicode MS";

        // Insert text that contains a discretionary ligature character (ﬁ).
        // Unicode character U+FB01 is the ligature 'fi'.
        builder.Writeln("Testing discretionary ligature: \uFB01");

        // Prepare TIFF save options.
        ImageSaveOptions tiffOptions = new ImageSaveOptions(SaveFormat.Tiff)
        {
            // Render all pages into a single multi‑page TIFF.
            // The default PageSet renders all pages, so we omit the ambiguous constructor.
            // Use the MultiPageLayout for TIFF frames to ensure a multi‑page file.
            PageLayout = MultiPageLayout.TiffFrames()
        };

        // Define output file path.
        string outputPath = Path.Combine(Directory.GetCurrentDirectory(), "LigatureTest.tiff");

        // Save the document as a TIFF image.
        doc.Save(outputPath, tiffOptions);

        // Verify that the TIFF file was created and has content.
        if (!File.Exists(outputPath))
            throw new FileNotFoundException("TIFF file was not created.", outputPath);

        FileInfo info = new FileInfo(outputPath);
        if (info.Length == 0)
            throw new InvalidOperationException("TIFF file is empty.");

        // Simple confirmation output.
        Console.WriteLine($"TIFF file generated successfully at: {outputPath}");
        Console.WriteLine($"File size: {info.Length} bytes");
    }
}
