using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Saving;

public class Program
{
    public static void Main()
    {
        // Define paths for the source document and the output TIFF.
        string docPath = Path.Combine(Directory.GetCurrentDirectory(), "Sample.docx");
        string tiffPath = Path.Combine(Directory.GetCurrentDirectory(), "FirstPage.tiff");

        // -----------------------------------------------------------------
        // 1. Create a multi‑page DOCX document.
        // -----------------------------------------------------------------
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Add three pages with simple text.
        builder.Writeln("This is page 1.");
        builder.InsertBreak(BreakType.PageBreak);
        builder.Writeln("This is page 2.");
        builder.InsertBreak(BreakType.PageBreak);
        builder.Writeln("This is page 3.");

        // Save the source document (optional, just to have a physical file).
        doc.Save(docPath);

        // -----------------------------------------------------------------
        // 2. Render only the first page to a single‑page TIFF.
        // -----------------------------------------------------------------
        ImageSaveOptions options = new ImageSaveOptions(SaveFormat.Tiff)
        {
            // Render only the first page (zero‑based index).
            PageSet = new PageSet(0)
        };

        // Save the TIFF.
        doc.Save(tiffPath, options);

        // -----------------------------------------------------------------
        // 3. Validate that the TIFF file was created.
        // -----------------------------------------------------------------
        if (!File.Exists(tiffPath))
        {
            throw new InvalidOperationException($"Failed to create TIFF file at '{tiffPath}'.");
        }

        // Output the result paths for confirmation (no interactive prompts).
        Console.WriteLine($"Source DOCX saved to: {docPath}");
        Console.WriteLine($"First page TIFF saved to: {tiffPath}");
    }
}
