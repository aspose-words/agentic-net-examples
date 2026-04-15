using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Saving;

public class Program
{
    public static void Main()
    {
        // Prepare output folder and file name.
        string outputDir = Path.Combine(Directory.GetCurrentDirectory(), "Output");
        Directory.CreateDirectory(outputDir);
        string tiffPath = Path.Combine(outputDir, "Multipage.tiff");

        // Create a blank document and add content spanning several pages.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        builder.Writeln("Page 1 - Hello Aspose.Words!");
        builder.InsertBreak(BreakType.PageBreak);
        builder.Writeln("Page 2 - Rendering to TIFF.");
        builder.InsertBreak(BreakType.PageBreak);
        builder.Writeln("Page 3 - End of document.");

        // Configure ImageSaveOptions for a multipage TIFF.
        ImageSaveOptions saveOptions = new ImageSaveOptions(SaveFormat.Tiff);
        // Use the layout that renders each page as a separate frame in the TIFF.
        saveOptions.PageLayout = MultiPageLayout.TiffFrames();

        // Save the document as a multipage TIFF.
        doc.Save(tiffPath, saveOptions);

        // Validate that the TIFF file was created.
        if (!File.Exists(tiffPath))
            throw new InvalidOperationException("Failed to create the TIFF file.");

        Console.WriteLine($"Multipage TIFF saved to: {tiffPath}");
    }
}
