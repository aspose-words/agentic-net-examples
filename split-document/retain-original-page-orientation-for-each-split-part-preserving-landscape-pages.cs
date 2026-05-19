using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Rendering;

public class SplitDocumentPreserveOrientation
{
    public static void Main()
    {
        // Prepare output directory.
        string outputDir = Path.Combine(Directory.GetCurrentDirectory(), "Output");
        Directory.CreateDirectory(outputDir);

        // Create a sample document with mixed page orientations.
        Document sourceDoc = new Document();
        DocumentBuilder builder = new DocumentBuilder(sourceDoc);

        // First page – default portrait orientation.
        builder.Writeln("This is page 1 (Portrait).");
        builder.InsertBreak(BreakType.PageBreak);

        // Insert a new section and set its orientation to Landscape.
        builder.InsertBreak(BreakType.SectionBreakNewPage);
        builder.PageSetup.Orientation = Orientation.Landscape;
        builder.Writeln("This is page 2 (Landscape).");
        builder.InsertBreak(BreakType.PageBreak);

        // Add a third page in the same landscape section.
        builder.Writeln("This is page 3 (Landscape).");

        // Save the source document (optional, useful for inspection).
        string sourcePath = Path.Combine(outputDir, "Sample.docx");
        sourceDoc.Save(sourcePath);

        // Split the document page by page, preserving the original orientation.
        int pageCount = sourceDoc.PageCount;
        for (int i = 0; i < pageCount; i++)
        {
            // Extract a single page. ExtractPages uses zero‑based index and a count of pages.
            Document part = sourceDoc.ExtractPages(i, 1);

            // Determine orientation of the extracted page for naming/debugging.
            PageInfo pageInfo = sourceDoc.GetPageInfo(i);
            string orientation = pageInfo.Landscape ? "Landscape" : "Portrait";

            // Save the split part.
            string partPath = Path.Combine(outputDir, $"Split_Page_{i + 1}_{orientation}.docx");
            part.Save(partPath);

            // Verify that the file was created.
            if (!File.Exists(partPath))
                throw new InvalidOperationException($"Failed to create split file: {partPath}");
        }

        Console.WriteLine($"Document split into {pageCount} parts. Files are located in: {outputDir}");
    }
}
