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
        string[] rangeSpecs = { "1-3", "5-7" };

        for (int i = 0; i < rangeSpecs.Length; i++)
        {
            // Parse the start and end page numbers.
            string[] parts = rangeSpecs[i].Split('-');
            int startPage = int.Parse(parts[0]); // 1‑based
            int endPage = int.Parse(parts[1]);   // 1‑based

            // Convert to zero‑based index for ExtractPages.
            int zeroBasedStart = startPage - 1;
            int pageCount = endPage - startPage + 1;

            // Extract the specified page range.
            Document extracted = sourceDoc.ExtractPages(zeroBasedStart, pageCount);

            // Save the extracted range as a PDF file.
            string outFile = Path.Combine(outputDir, $"Part_{i + 1}_{startPage}_to_{endPage}.pdf");
            extracted.Save(outFile, SaveFormat.Pdf);

            // Verify that the file was created.
            if (!File.Exists(outFile))
                throw new InvalidOperationException($"Failed to create split PDF: {outFile}");
        }
    }
}
