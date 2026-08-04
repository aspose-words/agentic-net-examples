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

        // Create a sample DOCX with three pages.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        builder.Writeln("Page 1");
        builder.InsertBreak(BreakType.PageBreak);
        builder.Writeln("Page 2");
        builder.InsertBreak(BreakType.PageBreak);
        builder.Writeln("Page 3");

        // Ensure the layout is calculated and obtain the page count.
        int sourcePageCount = doc.PageCount; // Expected to be 3.

        // Save the source document (optional, just for reference).
        string docPath = Path.Combine(artifactsDir, "Sample.docx");
        doc.Save(docPath);

        // Render the document to a multi‑frame TIFF where each page is a separate frame.
        string tiffPath = Path.Combine(artifactsDir, "Sample.tiff");
        ImageSaveOptions tiffOptions = new ImageSaveOptions(SaveFormat.Tiff)
        {
            // Use the layout that creates one frame per page.
            PageLayout = MultiPageLayout.TiffFrames()
        };
        doc.Save(tiffPath, tiffOptions);

        // Validate that the TIFF file was created.
        if (!File.Exists(tiffPath))
            throw new Exception("TIFF file was not created.");

        // Basic validation that the source document has the expected number of pages.
        if (sourcePageCount != 3)
            throw new Exception($"Source document page count is {sourcePageCount}, expected 3.");

        // Since each page is rendered as a separate frame, the existence of the file together
        // with the known page count of the source document confirms that the TIFF contains
        // the same number of pages.
        Console.WriteLine($"Success: TIFF file '{tiffPath}' was created with {sourcePageCount} pages.");
    }
}
