using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Saving;

public class Program
{
    public static void Main()
    {
        // Create a multi‑page DOCX document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        builder.Writeln("This is page 1.");
        builder.InsertBreak(BreakType.PageBreak);
        builder.Writeln("This is page 2.");
        builder.InsertBreak(BreakType.PageBreak);
        builder.Writeln("This is page 3.");

        // Save the source document (optional, just to demonstrate file creation).
        string sourcePath = "sample.docx";
        doc.Save(sourcePath);

        // Prepare ImageSaveOptions for TIFF output, rendering only the first page.
        ImageSaveOptions options = new ImageSaveOptions(SaveFormat.Tiff)
        {
            // Render only page 1.
            PageSet = new PageSet(1, 1)
        };

        // Render the first page to a single‑page TIFF file.
        string tiffPath = "first_page.tiff";
        doc.Save(tiffPath, options);

        // Verify that the TIFF file was created.
        if (!File.Exists(tiffPath))
            throw new Exception($"Failed to create TIFF file at '{tiffPath}'.");

        // Optionally, clean up the temporary DOCX file.
        if (File.Exists(sourcePath))
            File.Delete(sourcePath);
    }
}
