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

        // Add three pages of content.
        builder.Writeln("Page 1.");
        builder.InsertBreak(BreakType.PageBreak);
        builder.Writeln("Page 2.");
        builder.InsertBreak(BreakType.PageBreak);
        builder.Writeln("Page 3.");

        // Configure image save options for a multipage TIFF.
        ImageSaveOptions options = new ImageSaveOptions(SaveFormat.Tiff)
        {
            // Use the TiffFrames layout to store each page as a separate frame.
            PageLayout = MultiPageLayout.TiffFrames(),
            // Optional: set resolution (dpi) for the rendered images.
            Resolution = 300
        };

        // Define the output file path.
        string outputPath = Path.Combine(Directory.GetCurrentDirectory(), "MultipageOutput.tiff");

        // Save the document as a multipage TIFF.
        doc.Save(outputPath, options);

        // Verify that the file was created.
        if (!File.Exists(outputPath))
            throw new InvalidOperationException("Failed to create the TIFF file.");

        Console.WriteLine($"Multipage TIFF saved successfully to: {outputPath}");
    }
}
