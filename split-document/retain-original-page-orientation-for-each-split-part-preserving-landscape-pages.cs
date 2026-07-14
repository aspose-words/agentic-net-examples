using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Rendering;

public class SplitDocumentPreserveOrientation
{
    public static void Main()
    {
        // Prepare output folder.
        string outputDir = Path.Combine(Directory.GetCurrentDirectory(), "Output");
        if (Directory.Exists(outputDir))
            Directory.Delete(outputDir, true);
        Directory.CreateDirectory(outputDir);

        // Create a sample document with mixed page orientations.
        Document sourceDoc = new Document();
        DocumentBuilder builder = new DocumentBuilder(sourceDoc);

        // First two pages – portrait orientation (default).
        builder.Writeln("Portrait page 1");
        builder.InsertBreak(BreakType.PageBreak);
        builder.Writeln("Portrait page 2");

        // Start a new section and switch to landscape orientation.
        builder.InsertBreak(BreakType.SectionBreakNewPage);
        builder.PageSetup.Orientation = Orientation.Landscape;
        builder.Writeln("Landscape page 3");
        builder.InsertBreak(BreakType.PageBreak);
        builder.Writeln("Landscape page 4");

        // Ensure the document layout is up‑to‑date.
        sourceDoc.UpdatePageLayout();

        // Split the document page by page, preserving each page's orientation.
        int pageCount = sourceDoc.PageCount;
        for (int i = 1; i <= pageCount; i++)
        {
            // Extract a single page. ExtractPages(startPageIndex, pageCount) uses zero‑based index.
            Document pageDoc = sourceDoc.ExtractPages(i - 1, 1);

            // Save the extracted page.
            string pagePath = Path.Combine(outputDir, $"Page_{i}.docx");
            pageDoc.Save(pagePath);
        }

        // Validate that each split part retains the original orientation.
        for (int i = 1; i <= pageCount; i++)
        {
            string pagePath = Path.Combine(outputDir, $"Page_{i}.docx");
            if (!File.Exists(pagePath))
                throw new FileNotFoundException($"Expected split file not found: {pagePath}");

            // Load the split document.
            Document splitDoc = new Document(pagePath);

            // Determine expected orientation from the original document.
            Orientation expected = sourceDoc.GetPageInfo(i - 1).Landscape
                ? Orientation.Landscape
                : Orientation.Portrait;

            // The extracted document contains a single section.
            Orientation actual = splitDoc.Sections[0].PageSetup.Orientation;

            if (actual != expected)
                throw new InvalidOperationException($"Orientation mismatch on page {i}: expected {expected}, got {actual}");
        }

        Console.WriteLine($"Document split into {pageCount} parts. Files are located in: {outputDir}");
    }
}
