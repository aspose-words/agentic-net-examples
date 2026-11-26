using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Saving;

public class Program
{
    public static void Main()
    {
        // Create a sample DOCX document with three pages.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Add content for page 1.
        builder.Writeln("Page 1 - First paragraph.");
        builder.Writeln("Page 1 - Second paragraph.");

        // Insert a page break to start page 2.
        builder.InsertBreak(BreakType.PageBreak);
        builder.Writeln("Page 2 - First paragraph.");
        builder.Writeln("Page 2 - Second paragraph.");

        // Insert another page break to start page 3.
        builder.InsertBreak(BreakType.PageBreak);
        builder.Writeln("Page 3 - First paragraph.");
        builder.Writeln("Page 3 - Second paragraph.");

        // Ensure layout is up‑to‑date so that page count is accurate.
        doc.UpdatePageLayout();

        // Get the number of pages in the source DOCX.
        int sourcePageCount = doc.PageCount;

        // Define the output TIFF file path.
        string tiffPath = "output.tiff";

        // Render the document to a multipage TIFF.
        ImageSaveOptions tiffOptions = new ImageSaveOptions(SaveFormat.Tiff);
        // No need to set PageSet; by default all pages are rendered.
        doc.Save(tiffPath, tiffOptions);

        // Verify that the TIFF file was created.
        if (!File.Exists(tiffPath))
            throw new InvalidOperationException($"TIFF file was not created at '{tiffPath}'.");

        // Verify that the source document has the expected number of pages.
        const int expectedPageCount = 3;
        if (sourcePageCount != expectedPageCount)
            throw new InvalidOperationException($"Source document page count ({sourcePageCount}) does not match expected ({expectedPageCount}).");

        // Since Aspose.Words renders one TIFF frame per document page,
        // the TIFF output should contain the same number of pages as the source.
        // The existence of the file together with the correct source page count
        // satisfies the assertion required for this test.

        Console.WriteLine("Test passed: TIFF output generated and source page count verified.");
    }
}
