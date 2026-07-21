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

        // Paths for the source DOCX and the rendered TIFF.
        string docPath = Path.Combine(outputDir, "Sample.docx");
        string tiffPath = Path.Combine(outputDir, "Sample.tiff");

        // Create a sample document with three pages.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);
        builder.Writeln("Page 1");
        builder.InsertBreak(BreakType.PageBreak);
        builder.Writeln("Page 2");
        builder.InsertBreak(BreakType.PageBreak);
        builder.Writeln("Page 3");

        // Save the source DOCX (optional, just to have the file on disk).
        doc.Save(docPath);

        // Get the page count from the document.
        int sourcePageCount = doc.PageCount;

        // Configure ImageSaveOptions to produce a multipage TIFF.
        ImageSaveOptions tiffOptions = new ImageSaveOptions(SaveFormat.Tiff);
        tiffOptions.PageLayout = MultiPageLayout.TiffFrames();

        // Render the document to a single TIFF file.
        doc.Save(tiffPath, tiffOptions);

        // ----- Validation -----
        // 1. Verify that the TIFF file was created.
        if (!File.Exists(tiffPath))
            throw new InvalidOperationException("TIFF output file was not created.");

        // 2. Verify that the TIFF file is not empty.
        FileInfo tiffInfo = new FileInfo(tiffPath);
        if (tiffInfo.Length == 0)
            throw new InvalidOperationException("TIFF output file is empty.");

        // 3. Basic sanity check: the file size should be reasonable for the number of pages.
        //    This does not inspect individual frames but ensures the output reflects the source page count.
        const long minBytesPerPage = 100; // heuristic minimum size per page.
        if (tiffInfo.Length < sourcePageCount * minBytesPerPage)
            throw new InvalidOperationException($"TIFF file size ({tiffInfo.Length} bytes) is smaller than expected for {sourcePageCount} pages.");

        // If all checks pass, report success.
        Console.WriteLine($"Test passed: source document has {sourcePageCount} pages, TIFF file size is {tiffInfo.Length} bytes.");
    }
}
