using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Saving;

public class Program
{
    public static void Main()
    {
        // Create a temporary folder for output files.
        string outputDir = Path.Combine(Directory.GetCurrentDirectory(), "Output");
        Directory.CreateDirectory(outputDir);

        // Build a sample multi‑page document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        builder.Writeln("First page.");
        builder.InsertBreak(BreakType.PageBreak);
        builder.Writeln("Second page.");
        builder.InsertBreak(BreakType.PageBreak);
        builder.Writeln("Third page.");

        // Configure image save options to render only the first page as a TIFF.
        ImageSaveOptions options = new ImageSaveOptions(SaveFormat.Tiff);
        options.PageSet = new PageSet(0); // Zero‑based index of the first page.

        // Save the rendered TIFF.
        string tiffPath = Path.Combine(outputDir, "FirstPage.tiff");
        doc.Save(tiffPath, options);

        // Verify that the file was created.
        if (!File.Exists(tiffPath))
            throw new FileNotFoundException("The TIFF file was not created.", tiffPath);
    }
}
