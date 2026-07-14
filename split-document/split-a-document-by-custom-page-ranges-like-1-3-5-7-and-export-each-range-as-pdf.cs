using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Saving;

public class SplitDocumentByCustomRanges
{
    public static void Main()
    {
        // Prepare output directory.
        string outputDir = Path.Combine(Directory.GetCurrentDirectory(), "Output");
        Directory.CreateDirectory(outputDir);

        // Create a sample document with at least 7 pages.
        Document sourceDoc = new Document();
        DocumentBuilder builder = new DocumentBuilder(sourceDoc);

        for (int i = 1; i <= 7; i++)
        {
            builder.Writeln($"This is page {i}.");
            if (i < 7)
                builder.InsertBreak(BreakType.PageBreak);
        }

        // Define the custom page ranges (1‑based inclusive).
        int[][] ranges = new int[][]
        {
            new int[] { 1, 3 }, // pages 1 to 3
            new int[] { 5, 7 }  // pages 5 to 7
        };

        // Process each range.
        foreach (var range in ranges)
        {
            int startPage = range[0];
            int endPage = range[1];

            // Validate range.
            if (startPage < 1 || endPage > sourceDoc.PageCount || startPage > endPage)
                throw new ArgumentException($"Invalid page range: {startPage}-{endPage}");

            // Convert to zero‑based index and calculate count.
            int zeroBasedStart = startPage - 1;
            int count = endPage - startPage + 1;

            // Extract the pages into a new document.
            Document extracted = sourceDoc.ExtractPages(zeroBasedStart, count);

            // Save the extracted document as PDF.
            string outFile = Path.Combine(outputDir, $"Extracted_{startPage}_{endPage}.pdf");
            extracted.Save(outFile, SaveFormat.Pdf);

            // Verify that the file was created.
            if (!File.Exists(outFile))
                throw new InvalidOperationException($"Failed to create file: {outFile}");
        }
    }
}
