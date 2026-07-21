using System;
using System.IO;
using Aspose.Words;

public class Program
{
    public static void Main()
    {
        // Prepare folders.
        string artifactsDir = Path.Combine(Directory.GetCurrentDirectory(), "Artifacts");
        Directory.CreateDirectory(artifactsDir);
        string outputDir = Path.Combine(artifactsDir, "SplitOutputs");
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
        var ranges = new (int From, int To)[]
        {
            (1, 3),
            (5, 7)
        };

        // Split each range and save as a separate PDF.
        foreach (var range in ranges)
        {
            // Convert to zero‑based index for ExtractPages.
            int startIndex = range.From - 1;
            int pageCount = range.To - range.From + 1;

            // Guard against invalid ranges.
            if (startIndex < 0 || startIndex + pageCount > sourceDoc.PageCount)
                throw new ArgumentOutOfRangeException($"Range {range.From}-{range.To} is outside the document page count.");

            Document part = sourceDoc.ExtractPages(startIndex, pageCount);
            string outPath = Path.Combine(outputDir, $"Part_{range.From}_{range.To}.pdf");
            part.Save(outPath, SaveFormat.Pdf);

            // Verify that the file was created.
            if (!File.Exists(outPath))
                throw new InvalidOperationException($"Failed to create split PDF: {outPath}");
        }
    }
}
