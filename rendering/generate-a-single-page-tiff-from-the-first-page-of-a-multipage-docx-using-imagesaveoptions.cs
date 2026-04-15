using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Saving;

public class Program
{
    public static void Main()
    {
        // Create a multi‑page document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        builder.Writeln("First page content.");
        builder.InsertBreak(BreakType.PageBreak);
        builder.Writeln("Second page content.");
        builder.InsertBreak(BreakType.PageBreak);
        builder.Writeln("Third page content.");

        // Configure image save options to render only the first page as TIFF.
        ImageSaveOptions options = new ImageSaveOptions(SaveFormat.Tiff)
        {
            PageSet = new PageSet(0), // zero‑based index of the first page
            Resolution = 300          // optional: set DPI
        };

        // Define output file path.
        string outputPath = Path.Combine(Directory.GetCurrentDirectory(), "FirstPage.tiff");

        // Save the first page as a single‑page TIFF.
        doc.Save(outputPath, options);

        // Verify that the file was created.
        if (!File.Exists(outputPath))
            throw new InvalidOperationException("The TIFF file was not created.");

        // Optionally, you could perform further checks here (e.g., file size).
    }
}
