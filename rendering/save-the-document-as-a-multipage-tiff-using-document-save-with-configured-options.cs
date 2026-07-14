using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Saving;

public class Program
{
    public static void Main()
    {
        // Define output folder and ensure it exists.
        string outputDir = Path.Combine(Directory.GetCurrentDirectory(), "Output");
        Directory.CreateDirectory(outputDir);

        // Create a sample document with several pages.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        builder.Writeln("First page.");
        builder.InsertBreak(BreakType.PageBreak);
        builder.Writeln("Second page.");
        builder.InsertBreak(BreakType.PageBreak);
        builder.Writeln("Third page.");

        // Configure image save options for a multipage TIFF.
        ImageSaveOptions saveOptions = new ImageSaveOptions(SaveFormat.Tiff)
        {
            // Render each page as a separate frame in the TIFF.
            PageLayout = MultiPageLayout.TiffFrames(),
            // Optional: set resolution (dpi) and compression.
            Resolution = 300,
            TiffCompression = TiffCompression.Lzw
        };

        // Save the document as a multipage TIFF.
        string tiffPath = Path.Combine(outputDir, "Document.MultiPage.tiff");
        doc.Save(tiffPath, saveOptions);

        // Verify that the file was created.
        if (!File.Exists(tiffPath))
            throw new InvalidOperationException("Failed to create the TIFF file.");

        // Indicate successful completion.
        Console.WriteLine("Multipage TIFF saved to: " + tiffPath);
    }
}
