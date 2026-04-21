using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Saving;

public class SplitDocumentByPageRanges
{
    public static void Main()
    {
        // Prepare output directory.
        string outputDir = Path.Combine(Directory.GetCurrentDirectory(), "Output");
        Directory.CreateDirectory(outputDir);

        // Create a sample document with 7 pages.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);
        for (int i = 1; i <= 7; i++)
        {
            builder.Writeln($"This is page {i}.");
            if (i < 7)
                builder.InsertBreak(BreakType.PageBreak);
        }

        // Define the custom page ranges (1‑3 and 5‑7) using 1‑based page numbers.
        var pageRanges = new (int From, int To)[]
        {
            (1, 3),
            (5, 7)
        };

        // Export each range as a separate PDF.
        foreach (var range in pageRanges)
        {
            // Convert to zero‑based indices required by PageRange.
            int startZero = range.From - 1;
            int endZero = range.To - 1;

            // Configure PDF save options with the specific page range.
            PdfSaveOptions pdfOptions = new PdfSaveOptions
            {
                PageSet = new PageSet(new PageRange(startZero, endZero))
            };

            string outFile = Path.Combine(outputDir, $"Pages_{range.From}_to_{range.To}.pdf");
            doc.Save(outFile, pdfOptions);

            // Validate that the file was created.
            if (!File.Exists(outFile))
                throw new InvalidOperationException($"Failed to create PDF for pages {range.From}-{range.To}.");
        }

        // Optional: indicate completion.
        Console.WriteLine("Document split and PDF export completed successfully.");
    }
}
