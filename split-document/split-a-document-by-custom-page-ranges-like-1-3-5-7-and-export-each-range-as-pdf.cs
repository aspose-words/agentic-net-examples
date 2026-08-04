using System;
using System.Collections.Generic;
using System.IO;
using Aspose.Words;
using Aspose.Words.Saving;

public class SplitDocumentByCustomPageRanges
{
    public static void Main()
    {
        // Prepare output directory.
        string outputDir = Path.Combine(Directory.GetCurrentDirectory(), "Output");
        Directory.CreateDirectory(outputDir);

        // Create a sample document with 7 pages.
        Document sourceDoc = new Document();
        DocumentBuilder builder = new DocumentBuilder(sourceDoc);
        for (int i = 1; i <= 7; i++)
        {
            builder.Writeln($"This is page {i}.");
            if (i < 7)
                builder.InsertBreak(BreakType.PageBreak);
        }

        // Define custom page ranges (1‑based inclusive).
        var pageRanges = new List<(int Start, int End)>
        {
            (1, 3),
            (5, 7)
        };

        // Process each range: extract pages and save as a separate PDF.
        foreach (var range in pageRanges)
        {
            int startIndex = range.Start - 1;                     // zero‑based start page
            int pageCount = range.End - range.Start + 1;          // number of pages to extract

            // Ensure the requested range is within the document bounds.
            if (startIndex < 0 || startIndex + pageCount > sourceDoc.PageCount)
                throw new ArgumentOutOfRangeException($"Range {range.Start}-{range.End} is outside the document page count.");

            // Extract the specified pages into a new document.
            Document extracted = sourceDoc.ExtractPages(startIndex, pageCount);

            // Save the extracted document as PDF.
            string outFile = Path.Combine(outputDir, $"Pages_{range.Start}_to_{range.End}.pdf");
            extracted.Save(outFile, SaveFormat.Pdf);

            // Validate that the file was created.
            if (!File.Exists(outFile))
                throw new InvalidOperationException($"Failed to create output file: {outFile}");
        }

        // Optional: indicate successful completion.
        Console.WriteLine("Document split completed. PDFs are located in: " + outputDir);
    }
}
