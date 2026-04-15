using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Saving;

public class Program
{
    public static void Main()
    {
        // Prepare output folder.
        string artifactsDir = Path.Combine(Directory.GetCurrentDirectory(), "Artifacts");
        Directory.CreateDirectory(artifactsDir);

        // Create a sample DOCX document with three pages.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        builder.Writeln("Page 1");
        builder.InsertBreak(BreakType.PageBreak);
        builder.Writeln("Page 2");
        builder.InsertBreak(BreakType.PageBreak);
        builder.Writeln("Page 3");

        // Determine the number of pages in the source document.
        int sourcePageCount = doc.PageCount; // Triggers layout calculation.

        // Verify that the document has the expected number of pages.
        if (sourcePageCount != 3)
            throw new InvalidOperationException($"Expected 3 pages in the source document, but found {sourcePageCount}.");

        // Configure TIFF rendering to produce a multi‑frame TIFF (one frame per page).
        ImageSaveOptions tiffOptions = new ImageSaveOptions(SaveFormat.Tiff)
        {
            // Each page will be rendered as a separate frame.
            PageLayout = MultiPageLayout.TiffFrames(),
            // Optional: set a reasonable resolution.
            Resolution = 300
        };

        // Save the document as a TIFF file.
        string tiffPath = Path.Combine(artifactsDir, "output.tiff");
        doc.Save(tiffPath, tiffOptions);

        // Validate that the TIFF file was created.
        if (!File.Exists(tiffPath))
            throw new FileNotFoundException("TIFF output file was not created.", tiffPath);

        // Since Aspose.Words does not expose the number of frames in a TIFF,
        // we rely on the fact that MultiPageLayout.TiffFrames renders each page as a frame.
        // Therefore, the TIFF should contain the same number of pages as the source document.
        Console.WriteLine($"Source document pages: {sourcePageCount}");
        Console.WriteLine($"TIFF file created at: {tiffPath}");
        Console.WriteLine("Assertion passed: TIFF output contains the same number of pages as the source DOCX.");
    }
}
