using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Saving;

public class Program
{
    public static void Main()
    {
        // Define output file path.
        string outputPath = Path.Combine(Directory.GetCurrentDirectory(), "Multipage.tiff");

        // Create a blank document and add several pages.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        builder.Writeln("Page 1.");
        builder.InsertBreak(BreakType.PageBreak);
        builder.Writeln("Page 2.");
        builder.InsertBreak(BreakType.PageBreak);
        builder.Writeln("Page 3.");

        // Configure image save options for a multipage TIFF.
        ImageSaveOptions options = new ImageSaveOptions(SaveFormat.Tiff)
        {
            // Each page will be saved as a separate frame in the TIFF file.
            PageLayout = MultiPageLayout.TiffFrames(),
            // Optional: set resolution (dpi) for the rendered images.
            Resolution = 300
        };

        // Save the document as a multipage TIFF.
        doc.Save(outputPath, options);

        // Verify that the file was created.
        if (!File.Exists(outputPath))
            throw new InvalidOperationException("Failed to create the TIFF file.");

        Console.WriteLine($"Multipage TIFF saved successfully to: {outputPath}");
    }
}
