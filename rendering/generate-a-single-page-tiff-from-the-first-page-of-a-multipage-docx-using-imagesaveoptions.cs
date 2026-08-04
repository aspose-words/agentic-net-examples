using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Saving;

public class Program
{
    public static void Main()
    {
        // Prepare output directory.
        string outputDir = Path.Combine(Directory.GetCurrentDirectory(), "Output");
        Directory.CreateDirectory(outputDir);

        // Create a sample multi‑page document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        builder.Writeln("Page 1 – first page.");
        builder.InsertBreak(BreakType.PageBreak);
        builder.Writeln("Page 2 – second page.");
        builder.InsertBreak(BreakType.PageBreak);
        builder.Writeln("Page 3 – third page.");

        // Configure image save options to render only the first page as a single‑page TIFF.
        ImageSaveOptions options = new ImageSaveOptions(SaveFormat.Tiff);
        options.PageSet = new PageSet(0);                     // Zero‑based index of the first page.
        options.PageLayout = MultiPageLayout.SinglePage();   // Ensure a single‑page output.

        // Save the first page as TIFF.
        string tiffPath = Path.Combine(outputDir, "FirstPage.tiff");
        doc.Save(tiffPath, options);

        // Verify that the file was created.
        if (!File.Exists(tiffPath))
            throw new InvalidOperationException("The TIFF file was not created.");

        Console.WriteLine($"TIFF file successfully created at: {tiffPath}");
    }
}
