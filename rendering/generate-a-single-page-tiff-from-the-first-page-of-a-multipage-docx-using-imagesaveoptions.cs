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

        // Create a multi‑page DOCX document.
        string docPath = Path.Combine(outputDir, "MultiPage.docx");
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        builder.Writeln("First page.");
        builder.InsertBreak(BreakType.PageBreak);
        builder.Writeln("Second page.");
        builder.InsertBreak(BreakType.PageBreak);
        builder.Writeln("Third page.");

        // Save the source document.
        doc.Save(docPath);

        // Load the document for rendering.
        Document loadedDoc = new Document(docPath);

        // Configure ImageSaveOptions to render only the first page to TIFF.
        ImageSaveOptions options = new ImageSaveOptions(SaveFormat.Tiff);
        options.PageSet = new PageSet(0); // zero‑based index of the first page.

        // Save the first page as a single‑page TIFF.
        string tiffPath = Path.Combine(outputDir, "FirstPage.tiff");
        loadedDoc.Save(tiffPath, options);

        // Verify that the TIFF file was created.
        if (!File.Exists(tiffPath))
            throw new InvalidOperationException("The TIFF file was not created.");

        Console.WriteLine("First page rendered to TIFF successfully at:");
        Console.WriteLine(tiffPath);
    }
}
