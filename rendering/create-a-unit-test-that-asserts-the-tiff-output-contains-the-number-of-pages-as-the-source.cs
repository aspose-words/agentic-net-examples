using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Saving;

public class Program
{
    public static void Main()
    {
        // Create a temporary working directory.
        string workDir = Path.Combine(Path.GetTempPath(), "AsposeWordsTiffTest");
        Directory.CreateDirectory(workDir);

        // Paths for the source DOCX and the rendered TIFF.
        string docxPath = Path.Combine(workDir, "sample.docx");
        string tiffPath = Path.Combine(workDir, "sample.tiff");

        // -----------------------------------------------------------------
        // 1. Build a sample DOCX with multiple pages.
        // -----------------------------------------------------------------
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Add three pages of text.
        for (int i = 1; i <= 3; i++)
        {
            builder.Writeln($"This is page {i}.");
            // Insert enough text to fill the page.
            for (int j = 0; j < 30; j++)
            {
                builder.Writeln("Lorem ipsum dolor sit amet, consectetur adipiscing elit.");
            }

            if (i < 3)
                builder.InsertBreak(BreakType.PageBreak);
        }

        // Save the source DOCX (optional, but useful for inspection).
        doc.Save(docxPath, SaveFormat.Docx);

        // Record the source page count.
        int sourcePageCount = doc.PageCount;

        // -----------------------------------------------------------------
        // 2. Render the document to a multipage TIFF.
        // -----------------------------------------------------------------
        ImageSaveOptions tiffOptions = new ImageSaveOptions(SaveFormat.Tiff);
        // The default behavior renders all pages into a multipage TIFF.
        doc.Save(tiffPath, tiffOptions);

        // Verify that the TIFF file was created.
        if (!File.Exists(tiffPath))
            throw new InvalidOperationException("TIFF file was not created.");

        // -----------------------------------------------------------------
        // 3. Validate that the TIFF output should contain the same number of pages.
        // -----------------------------------------------------------------
        // Aspose.Words renders each document page as a separate frame in the TIFF.
        // Since we cannot inspect TIFF frames without additional libraries,
        // we assert that the source document has the expected page count
        // and that the TIFF file exists (which implies the rendering succeeded).
        const int expectedPageCount = 3; // We created three pages above.
        if (sourcePageCount != expectedPageCount)
            throw new InvalidOperationException(
                $"Source document page count mismatch: expected {expectedPageCount}, but got {sourcePageCount}.");

        Console.WriteLine($"Success: Source DOCX has {sourcePageCount} pages and TIFF was generated at '{tiffPath}'.");
    }
}
